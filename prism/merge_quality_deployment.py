"""Merge the three fixed whole-quality adapters into ordinary Backbonecs weights."""
from pathlib import Path
import hashlib
import json
import shutil
import subprocess
import sys
import torch
import numpy as np
from PIL import Image
from torchvision.transforms import Compose, ToTensor, Normalize
from probe_adapter_strength import load_official, load_model, restored_image

R = Path('/root/autodl-tmp/bilalora')
F = R / 'quality_selection'


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


@torch.inference_mode()
def main():
    assert (F / 'logs/final.exit').read_text().strip() == '0'
    assert (F / 'logs/mechanism.exit').read_text().strip() == '0'
    torch.set_num_threads(4)
    official = load_official()
    destination = F / 'deployment'
    destination.mkdir(exist_ok=True)
    shutil.copytree(R / 'official/DEAmodel', destination / 'DEAmodel', dirs_exist_ok=True,
                    ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
    shutil.copy2(R / 'additional_baselines/deploy_quality_model.py', destination / 'dehaze.py')
    transform = Compose([ToTensor(), Normalize((.48145466, .4578275, .40821073), (.26862954, .26130258, .27577711))])
    inputs = json.loads((F / 'inputs_manifest.json').read_text())
    input_path = Path(inputs['scopes']['RTTS128']['input_dir']) / inputs['scopes']['RTTS128']['images'][0]['name']
    smoke = destination / 'smoke_input'
    smoke.mkdir(exist_ok=True)
    shutil.copy2(input_path, smoke / input_path.name)
    expected = destination / 'smoke_expected'
    expected.mkdir(exist_ok=True)
    records = []
    for seed in [3407, 3408, 3409]:
        source = R / f'quality_repair_fullimage/sum_seed{seed}/epoch_5_lora_adapter_finetuned.pth'
        model, layers, scales = load_model(official, source)
        _, raw_before, png_before = restored_image(model, input_path, transform)
        count = 0
        for name, layer in list(model.named_modules()):
            if not isinstance(layer, official.LoRALayer):
                continue
            conv = layer.original_layer
            weight = conv.weight + torch.sigmoid(layer.alpha) * layer.scaling * (layer.lora_B @ layer.lora_A).view_as(conv.weight)
            conv.weight.copy_(weight)
            # Official LoRALayer uses direct conv2d with numeric padding.
            conv.padding_mode = 'zeros'
            parent, child = name.rsplit('.', 1)
            setattr(model.get_submodule(parent), child, conv)
            count += 1
        assert count == 6 and not any(isinstance(layer, official.LoRALayer) for layer in model.modules())
        assert set(model.state_dict()) == set(torch.load(R / 'official/weight/Base.pth', map_location='cpu', weights_only=True))
        _, raw_after, png_after = restored_image(model, input_path, transform)
        error = float((raw_after - raw_before).abs().max().cpu())
        pixel_error = int(((png_after - png_before).abs() * 255).max().round().cpu())
        assert error <= 2e-5 and pixel_error <= 1
        target = destination / f'dehaze_eva_fullquality_seed{seed}.pth'
        torch.save({name: value.cpu() for name, value in model.state_dict().items()}, target)
        records.append({'seed': seed, 'source_adapter_sha256': sha(source), 'merged_weight_sha256': sha(target),
                        'merged_parameters': sum(parameter.numel() for parameter in model.parameters()),
                        'merged_convolutions': count, 'smoke_input': input_path.name,
                        'raw_max_abs_error': error, 'PNG_max_LSB_error': pixel_error,
                        'PNG_exact': bool(torch.equal(png_before, png_after)),
                        'protocol': 'Fixed35 total epochs, coldEVA30 + wholeimageSUM5; seed not chosen by external scores'})
        array = (png_after[0].permute(1, 2, 0).cpu().numpy() * 255).round().astype(np.uint8)
        Image.fromarray(array).save(expected / f'seed{seed}.png')
        del model, layers
        torch.cuda.empty_cache()
    subprocess.run([sys.executable, str(destination / 'dehaze.py'), '--weights', str(destination / 'dehaze_eva_fullquality_seed3407.pth'),
                    '--input', str(smoke), '--output', str(destination / 'smoke_cli_output'), '--device', 'cuda'], check=True)
    cli = np.asarray(Image.open(destination / 'smoke_cli_output' / input_path.name).convert('RGB'))
    reference = np.asarray(Image.open(expected / 'seed3407.png').convert('RGB'))
    assert cli.shape == reference.shape
    cli_error = int(np.abs(cli.astype(np.int16) - reference.astype(np.int16)).max())
    assert cli_error <= 1
    records[0]['standalone_CLI_max_LSB_error'] = cli_error
    (destination / 'merge_audit.json').write_text(json.dumps(records, indent=2))
    (destination / 'README.md').write_text('''# Merged research checkpoints

Fixed seed3407 is the default demonstration, with3408/3409 retained as repeats.
Training: same published Base, EVA initialization, H2C-selected positions,30epochs, then5epochs of local H2C + whole-composition MUSIQ quality loss. No test-selected checkpoint.

The adapter is merged into the ordinary backbone. The inference script has no CLIP/MUSIQ dependency; it needs PyTorch, torchvision, Pillow and the bundled DEAmodel dependencies. It accepts --device cpu if no cloudGPU is available. No localGPU is invoked automatically in the research workflow.

Example on an explicitly chosen device:
python dehaze.py --weights dehaze_eva_fullquality_seed3407.pth --input YOUR_IMAGES --output RESTORED --device cuda

Outputs are PNG with the original stem. Keep input basenames unique. Exact official normalize-before-bicubic geometry and output rounding are retained. See merge_audit.json for the actual one-image merge check, not an all-image numerical guarantee.
''')
    (destination / 'requirements.txt').write_text('torch>=2.5\ntorchvision\nPillow>=9.1\nnumpy\neinops\n')
    print('QUALITY_DEPLOYMENT_MERGED', json.dumps(records), flush=True)


if __name__ == '__main__':
    main()
