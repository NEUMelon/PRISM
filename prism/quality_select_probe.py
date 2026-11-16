# -*- coding: utf-8 -*-
"""Probe quality-constrained layer selection using architecture-validation images.

Run on the cloud GPU. It reads no test images or paired clean references.
The 32-image selection is frozen before the separately filtered verification check.
"""
from pathlib import Path
import argparse
import hashlib
import importlib.metadata
import importlib.util
import itertools
import json
import math
import os
import shutil
import sys
import tempfile
import time

import numpy as np
from PIL import Image

R = Path('/root/autodl-tmp/bilalora')
F = R / 'quality_selection/probes'
BLOCKS = [f'level3_block{i}' for i in range(1, 9)]
ALL = [f'down_level{level}_block{i}' for level in [1, 2] for i in range(1, 5)] + BLOCKS
PUBLISHED = ('level3_block1', 'level3_block7', 'level3_block8')


def sha(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()


def subset_key(layers):
    return '|'.join(sorted(layers))


def decoded_hash(path):
    with Image.open(path) as source:
        image = source.convert('RGB')
        return hashlib.sha256(str(image.size).encode() + image.tobytes()).hexdigest()


def filter_verification(reference):
    """Remove decoded duplicates with the weight-training half before scoring."""
    import torch
    folder = R / 'data/500_foggy'
    names = sorted(path.name for path in folder.iterdir() if path.is_file())
    assert len(names) == 500
    permutation = torch.randperm(500, generator=torch.Generator().manual_seed(3407)).tolist()
    assert [row['name'] for row in reference['images']] == [names[index] for index in permutation[250:282]]
    assert [row['name'] for row in reference['verification_images']] == [names[index] for index in permutation[282:346]]
    trained = [{'name': names[index], 'decoded_pixel_sha256': decoded_hash(folder / names[index])}
               for index in permutation[:250]]
    trained_by_hash = {}
    for row in trained:
        trained_by_hash.setdefault(row['decoded_pixel_sha256'], []).append(row['name'])
    selection_hashes = {row['name']: decoded_hash(folder / row['name']) for row in reference['images']}
    assert len(set(selection_hashes.values())) == 32
    assert not set(selection_hashes.values()).intersection(trained_by_hash), 'Selection32 overlaps weight-training250'
    kept = []
    excluded = []
    verification_hashes = {}
    for record in reference['verification_images']:
        digest = decoded_hash(folder / record['name'])
        verification_hashes[record['name']] = digest
        if digest in trained_by_hash:
            excluded.append({'name': record['name'], 'decoded_pixel_sha256': digest,
                             'matching_weight_training_names': trained_by_hash[digest]})
        else:
            kept.append(record)
    assert len(kept) > 0
    assert not set(selection_hashes.values()).intersection(verification_hashes[row['name']] for row in kept)
    audit = {'weight_training_images': trained, 'selection32_decoded_hashes': selection_hashes,
             'selection32_weight_training_decoded_overlap': [],
             'original_verification64_names': [row['name'] for row in reference['verification_images']],
             'verification64_decoded_hashes': verification_hashes,
             'verification_excluded_weight_training_duplicates': excluded,
             'verification_kept_names': [row['name'] for row in kept], 'verification_kept_count': len(kept),
             'verification_unique_decoded_count': len({verification_hashes[row['name']] for row in kept}),
             'interpretation': 'Both selection and verification originate in the original architecture-learning half. The verification check is not an external test.'}
    return kept, audit


def select_quality_constrained(rows):
    """Minimize H2C among triples meeting the published triple's quality score."""
    assert len(rows) == 56
    assert {tuple(row['selected']) for row in rows} == set(itertools.combinations(BLOCKS, 3))
    assert all(math.isfinite(row['h2c32_loss']) and math.isfinite(row['musiq']) for row in rows)
    baseline = next(row for row in rows if tuple(row['selected']) == PUBLISHED)
    feasible = [row for row in rows if row['musiq'] >= baseline['musiq']]
    chosen = min(feasible, key=lambda row: (row['h2c32_loss'], row['selected']))
    quality_best = min(rows, key=lambda row: (-row['musiq'], row['h2c32_loss'], row['selected']))
    h2c_best = min(rows, key=lambda row: (row['h2c32_loss'], row['selected']))
    pareto = []
    for row in rows:
        dominated = any(other['h2c32_loss'] <= row['h2c32_loss'] and other['musiq'] >= row['musiq']
                        and (other['h2c32_loss'] < row['h2c32_loss'] or other['musiq'] > row['musiq'])
                        for other in rows)
        if not dominated:
            pareto.append(row['selected'])
    return {'quality_threshold': baseline['musiq'], 'feasible_triples': len(feasible),
            'published': baseline, 'quality_constrained': chosen, 'quality_best': quality_best,
            'h2c_best': h2c_best, 'pareto_frontier': pareto,
            'rule': 'Min selection32 H2C among full-image selection32 MUSIQ >= published triple MUSIQ; ties lexicographic.'}


def load_official():
    sys.path.insert(0, str(R / 'official'))
    spec = importlib.util.spec_from_file_location('quality_probe_official', R / 'official/BiLaLoRA_Test.py')
    official = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(official)
    return official


def build_model(official, base_state, layers, adapter_state):
    import torch
    model = official.Backbonecs().cuda()
    model.load_state_dict(base_state, strict=True)
    model, _ = official.inject_lora_to_layers(model, layers, 8, 16, 'cuda')
    loaded = model.load_state_dict(adapter_state, strict=False)
    expected = {name for name, _ in model.named_parameters() if 'lora_' in name or name.endswith('.alpha')}
    assert not loaded.unexpected_keys and expected == set(adapter_state)
    assert all(torch.equal(parameter.detach().cpu(), adapter_state[name].cpu())
               for name, parameter in model.named_parameters() if name in adapter_state)
    model.eval()
    for parameter in model.parameters():
        parameter.requires_grad = False
    return model


def prepare_inputs(records, max_edge):
    """Match official deployment, including normalization before bicubic resize."""
    from torchvision.transforms import ToTensor, Normalize, Resize, InterpolationMode
    normalize = Normalize((.48145466, .4578275, .40821073), (.26862954, .26130258, .27577711))
    inputs = []
    metadata = []
    for record in records:
        path = R / 'data/500_foggy' / record['name']
        assert sha(path) == record['file_sha256']
        with Image.open(path) as source:
            image = source.convert('RGB')
            source_size = image.size
            decoded = hashlib.sha256(str(image.size).encode() + image.tobytes()).hexdigest()
            if max_edge and max(image.size) > max_edge:
                scale = max_edge / max(image.size)
                image = image.resize((max(16, round(image.width * scale)), max(16, round(image.height * scale))), Image.Resampling.BICUBIC)
            raw = ToTensor()(image)
        height, width = raw.shape[1:]
        assert min(height, width) >= 16
        normalized = normalize(raw)
        inputs.append({'name': record['name'], 'tensor': normalized.unsqueeze(0), 'height': height, 'width': width})
        metadata.append({'name': record['name'], 'file_sha256': record['file_sha256'], 'split_index': record['split_index'],
                         'decoded_pixel_sha256': decoded,
                         'source_size': list(source_size), 'scored_size': [width, height],
                         'normalized_input_sha256': hashlib.sha256(normalized.numpy().tobytes()).hexdigest(),
                         'network_size': [width // 16 * 16, height // 16 * 16]})
    return inputs, metadata


def restored_uint8(model, item):
    import torch
    from torchvision.transforms import Resize, InterpolationMode
    incoming = item['tensor'].cuda()
    incoming = Resize((item['height'] // 16 * 16, item['width'] // 16 * 16),
                      interpolation=InterpolationMode.BICUBIC, antialias=True)(incoming)
    outgoing = model(incoming)
    if isinstance(outgoing, tuple):
        outgoing = outgoing[0]
    outgoing = Resize((item['height'], item['width']), interpolation=InterpolationMode.BICUBIC,
                      antialias=True)(outgoing.squeeze(0))
    # torchvision.save_image performs these operations before casting to uint8.
    array = outgoing.mul(255).add_(.5).clamp_(0, 255).permute(1, 2, 0).to('cpu', torch.uint8).numpy()
    return array


def score_model(model, metric, inputs, verify_file=False):
    import torch
    values = []
    check = None
    started = time.monotonic()
    for index, item in enumerate(inputs):
        array = restored_uint8(model, item)
        tensor = torch.from_numpy(array.copy()).permute(2, 0, 1).float().div(255).unsqueeze(0).cuda()
        score = float(metric(tensor).cpu())
        assert math.isfinite(score)
        values.append(score)
        if verify_file and index == 0:
            with tempfile.TemporaryDirectory(prefix='quality_probe_png_') as directory:
                path = Path(directory) / 'quantized.png'
                Image.fromarray(array).save(path)
                assert np.array_equal(np.asarray(Image.open(path)), array)
                file_score = float(metric(str(path)).cpu())
                error = abs(score - file_score)
                assert error <= 1e-4, (score, file_score, error)
                check = {'file_score': file_score, 'tensor_score': score, 'absolute_error': error,
                         'decoded_png_pixels_exact': True, 'name': item['name']}
        del tensor
    result = {'musiq': float(np.mean(values)), 'per_image': values, 'seconds': time.monotonic() - started}
    if check is not None:
        result['file_tensor_equivalence'] = check
    return result


def append_row(path, row):
    with path.open('a', encoding='utf-8') as stream:
        stream.write(json.dumps(row) + '\n')
        stream.flush()
        os.fsync(stream.fileno())


def read_rows(path):
    if not path.exists():
        return []
    rows = []
    lines = path.read_text().splitlines()
    for index, line in enumerate(lines):
        try:
            rows.append(json.loads(line))
        except json.JSONDecodeError:
            if index != len(lines) - 1:
                raise ValueError(f'Invalid interior JSONL row at {path}:{index + 1}')
            backup = path.with_name(path.stem + f'.interrupted_{time.time_ns()}.jsonl')
            shutil.copy2(path, backup)
            temporary = path.with_suffix('.recovery.tmp')
            temporary.write_text(''.join(json.dumps(row) + '\n' for row in rows), encoding='utf-8')
            temporary.replace(path)
            print('RECOVERED_PARTIAL_JSONL_ROW', str(backup), flush=True)
    return rows


def mask_model(model, selected, original_scales):
    for name in ALL:
        for layer in model.get_submodule(name).modules():
            if hasattr(layer, 'lora_A'):
                layer.scaling = original_scales[id(layer)] if name in selected else 0.0


def compare_final_models(official, base_state, mode, seed, inputs, metric, output):
    import torch
    runs = json.loads((R / 'selection_contribution/protocol_manifest.json').read_text())['runs']
    candidates = [run for run in runs if run['initializer'] == mode and run['seed'] == seed]
    assert len(candidates) == 3
    existing = read_rows(output)
    done = {row['position']: row for row in existing}
    assert len(done) == len(existing)
    for run in candidates:
        checkpoint_path = R / run['work'] / 'experiment/DEA/BiLaLoRA/saved_model/epoch_30_lora_adapter_finetuned.pth'
        digest = sha(checkpoint_path)
        if run['position'] in done:
            assert done[run['position']]['checkpoint_sha256'] == digest
            continue
        checkpoint = torch.load(checkpoint_path, map_location='cpu', weights_only=True)
        config = checkpoint.get('lora_config', checkpoint.get('nas_lora_config'))
        assert config['epoch'] == 30 and set(config['selected_layers']) == set(run['selected'])
        model = build_model(official, base_state, run['selected'], checkpoint['lora_state_dict'])
        row = {'position': run['position'], 'selected': sorted(run['selected']), 'checkpoint_sha256': digest,
               **score_model(model, metric, inputs, verify_file=not existing)}
        append_row(output, row)
        existing.append(row)
        print('FINAL30_QUALITY_PROBED', mode, seed, run['position'], row['musiq'], flush=True)
        del model
        torch.cuda.empty_cache()
    return existing


def main():
    import torch
    import pyiqa
    parser = argparse.ArgumentParser()
    parser.add_argument('--mode', choices=['baseline', 'eva_centered'], required=True)
    parser.add_argument('--seed', type=int, choices=[3407, 3408, 3409], required=True)
    parser.add_argument('--limit', type=int, help='Smoke execution only. Never writes a frozen selection or verifies held images.')
    parser.add_argument('--max-edge', type=int, default=0, help='Predeclared scoring cap. Zero matches full-size deployment.')
    parser.add_argument('--final-only', action='store_true', help='Score the three final30 models on selection32, without subset search.')
    parser.add_argument('--skip-final', action='store_true')
    args = parser.parse_args()
    assert torch.cuda.is_available(), 'Cloud GPU required; do not execute on the local GPU'
    assert args.limit is None or 0 < args.limit <= 56
    assert args.max_edge == 0 or args.max_edge >= 224
    torch.set_num_threads(4)
    suffix = 'full' if args.max_edge == 0 else f'maxedge{args.max_edge}'
    output = F / f'{args.mode}_{args.seed}' / suffix
    output.mkdir(parents=True, exist_ok=True)
    reference = json.loads((R / 'selection_diagnostic/protocol_manifest.json').read_text())
    diagnostic_path = R / f'selection_diagnostic/{args.mode}_{args.seed}.json'
    diagnostic = json.loads(diagnostic_path.read_text())
    verification_records_filtered, overlap_audit = filter_verification(reference)
    snapshot_path = R / f'experiments/selection_diagnostic_{args.mode}_seed{args.seed}/full_epoch5_search.pth'
    base_path = R / 'official/weight/Base.pth'
    metric_path = R / 'metrics/musiq_koniq_ckpt-e95806b9.pth'
    assert sha(metric_path) == 'e95806b9eae5f3814c410f574ba8e552362bd5bc63d758ed5b97860f5d6185aa'
    assert sha(snapshot_path) == diagnostic['checkpoint_sha256']
    inputs, records = prepare_inputs(reference['images'], args.max_edge)
    assert len(inputs) == 32
    protocol = {'mode': args.mode, 'seed': args.seed, 'selection_images': records,
                'verification_names': [record['name'] for record in verification_records_filtered],
                'overlap_audit': overlap_audit,
                'base_sha256': sha(base_path), 'snapshot_sha256': sha(snapshot_path),
                'diagnostic_sha256': sha(diagnostic_path), 'metric_sha256': sha(metric_path),
                'script_sha256': sha(Path(__file__)), 'official_inference_sha256': sha(R / 'official/BiLaLoRA_Test.py'),
                'torch': torch.__version__, 'pyiqa': importlib.metadata.version('pyiqa'), 'gpu': torch.cuda.get_device_name(0),
                'max_edge': args.max_edge, 'scoring': 'Official full-image Normalize, bicubic resize to16multiple, restore scored dimensions, save_image uint8 quantization, pinned MUSIQ on RGBfloat32decoded-equivalent tensor.',
                'h2c_source': 'Frozen centered224 selection32 losses from earlier full16 snapshot subset diagnostic',
                'selection_rule': 'Minimum H2C32 among quality32 >= published quality32; baseline always feasible.',
                'data_scope': 'Only architecture-learning half for selection32 and filtered verification after removing decoded overlap with weight-training250. No test image or clean reference read.',
                'smoke_policy': 'Any --limit invocation cannot freeze or evaluate verification images, even if cache already contains56 rows.'}
    protocol_path = output / 'protocol.json'
    if protocol_path.exists():
        assert json.loads(protocol_path.read_text()) == protocol, 'Existing cache belongs to a different protocol or code'
    else:
        protocol_path.write_text(json.dumps(protocol, indent=2))
    official = load_official()
    base_state = torch.load(base_path, map_location='cpu', weights_only=True)
    metric = pyiqa.create_metric('musiq', device='cuda', pretrained_model_path=str(metric_path))
    metric.eval()
    with torch.inference_mode():
        if not args.skip_final:
            compare_final_models(official, base_state, args.mode, args.seed, inputs, metric, output / 'final30_selection32.jsonl')
        if args.final_only:
            print('FINAL_ONLY_PROBE_COMPLETED', flush=True)
            return
        snapshot = torch.load(snapshot_path, map_location='cpu', weights_only=True)
        model = build_model(official, base_state, ALL, snapshot['lora_state_dict'])
        scales = {id(layer): layer.scaling for layer in model.modules() if hasattr(layer, 'lora_A')}
        h2c = {subset_key(row['selected']): row['loss'] for row in diagnostic['subsets']}
        path = output / 'subsets32.jsonl'
        existing = read_rows(path)
        done = {subset_key(row['selected']): row for row in existing}
        assert len(done) == len(existing)
        subsets = list(itertools.combinations(BLOCKS, 3))
        # Published first provides a baseline for a short smoke execution.
        subsets.sort(key=lambda group: (group != PUBLISHED, group))
        if args.limit is not None:
            subsets = subsets[:args.limit]
        for selected in subsets:
            key = subset_key(selected)
            if key in done:
                assert done[key]['h2c32_loss'] == h2c[key] and len(done[key]['per_image']) == 32
                continue
            mask_model(model, selected, scales)
            row = {'selected': list(selected), 'h2c32_loss': h2c[key],
                   **score_model(model, metric, inputs, verify_file=not existing)}
            append_row(path, row)
            existing.append(row)
            done[key] = row
            print('QUALITY_SUBSET_PROBED', args.mode, args.seed, selected, row['musiq'], flush=True)
        if args.limit is not None:
            print('SMOKE_COMPLETED_NO_SELECTION', len(existing), flush=True)
            return
        frozen = select_quality_constrained(existing)
        frozen['selection_protocol_sha256'] = sha(protocol_path)
        frozen['subsets32_sha256'] = sha(path)
        frozen_path = output / 'frozen_selection.json'
        if frozen_path.exists():
            assert json.loads(frozen_path.read_text()) == frozen
        else:
            frozen_path.write_text(json.dumps(frozen, indent=2))
        print('QUALITY_SELECTION_FROZEN', frozen['quality_constrained']['selected'], flush=True)
        verification_inputs, verification_records = prepare_inputs(verification_records_filtered, args.max_edge)
        (output / 'verification_inputs.json').write_text(json.dumps(verification_records, indent=2))
        rules = {'published': frozen['published']['selected'], 'quality_constrained': frozen['quality_constrained']['selected'],
                 'quality_best': frozen['quality_best']['selected'], 'h2c_best': frozen['h2c_best']['selected'],
                 'gate': diagnostic['gate_selected']}
        verification_path = output / 'verification.jsonl'
        verified = read_rows(verification_path)
        verified_by_rule = {row['rule']: row for row in verified}
        assert len(verified_by_rule) == len(verified)
        cache = {subset_key(row['selected']): {key: row[key] for key in ['musiq', 'per_image', 'seconds']} for row in verified}
        for rule, selected in rules.items():
            if rule in verified_by_rule:
                assert verified_by_rule[rule]['selected'] == selected
                continue
            key = subset_key(selected)
            if key not in cache:
                mask_model(model, selected, scales)
                cache[key] = score_model(model, metric, verification_inputs)
            row = {'rule': rule, 'selected': selected, **cache[key], 'frozen_selection_sha256': sha(frozen_path)}
            append_row(verification_path, row)
            print('QUALITY_RULE_VERIFIED', rule, len(verification_inputs), row['musiq'], flush=True)
    print('QUALITY_SELECTION_PROBE_COMPLETED', args.mode, args.seed, flush=True)


if __name__ == '__main__':
    main()
