"""Inference for a merged research checkpoint; device must be chosen explicitly."""
from pathlib import Path
import argparse
import json
import sys
import torch
import numpy as np
from PIL import Image
from torchvision.transforms import Compose, ToTensor, Normalize, Resize, InterpolationMode


@torch.inference_mode()
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--weights', type=Path, required=True)
    parser.add_argument('--input', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--device', choices=['cpu', 'cuda'], required=True)
    args = parser.parse_args()
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from DEAmodel import Backbonecs
    device = torch.device(args.device)
    if device.type == 'cuda':
        assert torch.cuda.is_available()
    torch.set_num_threads(4)
    model = Backbonecs().to(device).eval()
    model.load_state_dict(torch.load(args.weights, map_location='cpu', weights_only=True), strict=True)
    files = sorted(path for path in args.input.iterdir() if path.suffix.lower() in ('.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff'))
    assert files and len({path.stem for path in files}) == len(files), 'Input stems must be unique'
    args.output.mkdir(parents=True, exist_ok=True)
    transform = Compose([ToTensor(), Normalize((.48145466, .4578275, .40821073), (.26862954, .26130258, .27577711))])
    for path in files:
        with Image.open(path) as image:
            incoming = transform(image.convert('RGB')).unsqueeze(0).to(device)
        h, w = incoming.shape[-2:]
        assert min(h, w) >= 16
        resize = lambda tensor, size: Resize(size, interpolation=InterpolationMode.BICUBIC, antialias=True)(tensor)
        output = model(resize(incoming, (h // 16 * 16, w // 16 * 16)))
        if isinstance(output, tuple):
            output = output[0]
        raw = resize(output[0], (h, w))
        assert torch.isfinite(raw).all()
        array = (raw * 255 + .5).clamp(0, 255).to(torch.uint8).permute(1, 2, 0).cpu().numpy()
        target = args.output / (path.stem + '.png')
        Image.fromarray(array).save(target)
        print(json.dumps({'input': path.name, 'output': target.name, 'width': w, 'height': h}), flush=True)


if __name__ == '__main__':
    main()
