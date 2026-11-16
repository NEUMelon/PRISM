"""Score saved RGB images with the frozen KonIQ MUSIQ checkpoint."""
from pathlib import Path
import argparse,csv,hashlib,json

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--input',type=Path,required=True)
    p.add_argument('--output',type=Path,required=True,help='Output per-image CSV')
    p.add_argument('--device',choices=['cpu','cuda'],required=True)
    p.add_argument('--weights',type=Path,default=Path(__file__).resolve().parent/'musiq_koniq_ckpt-e95806b9.pth')
    args=p.parse_args()
    import torch,pyiqa
    torch.set_num_threads(4)
    assert args.weights.is_file(),args.weights
    metric=pyiqa.create_metric('musiq',device=args.device,pretrained_model_path=str(args.weights)).eval()
    images=sorted(x for x in args.input.iterdir() if x.suffix.lower() in {'.png','.jpg','.jpeg','.bmp','.tif','.tiff'})
    assert images,'No images found'
    args.output.parent.mkdir(parents=True,exist_ok=True)
    values=[]
    with args.output.open('w',encoding='utf-8',newline='') as f:
        writer=csv.DictWriter(f,fieldnames=['filename','MUSIQ']);writer.writeheader()
        for image in images:
            with torch.inference_mode():value=float(metric(str(image)).cpu())
            writer.writerow(dict(filename=image.name,MUSIQ=value));f.flush();values.append(value)
    info=dict(count=len(values),mean=sum(values)/len(values),device=args.device,
              weights_sha256=hashlib.sha256(args.weights.read_bytes()).hexdigest(),
              torch=torch.__version__,pyiqa=pyiqa.__version__)
    args.output.with_suffix('.json').write_text(json.dumps(info,indent=2),encoding='utf-8')
    print(json.dumps(info))

if __name__=='__main__':main()
