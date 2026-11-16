"""Fixed-rank convolutional adaptation of EVA (NeurIPS 2025, Sec.3.2).

Our exploratory variant uses paired re-hazing response moments instead of raw
activation moments. This is not the author's full adaptive-rank EVA algorithm.
"""
from pathlib import Path
import json,random,time
import numpy as np
import torch
import torch.nn.functional as F
from PIL import Image
from torchvision.transforms import Resize,CenterCrop,ToTensor,Normalize

@torch.no_grad()
def calibrate_lora(model,train_subset,mode):
    assert mode in ['random_ortho','eva','haze_delta','eva_centered','delta_centered','gcpca']
    start=time.perf_counter()
    rng=(random.getstate(),np.random.get_state(),torch.random.get_rng_state(),torch.cuda.get_rng_state_all())
    layers={n:m for n,m in model.named_modules() if hasattr(m,'lora_A') and hasattr(m,'original_layer')}
    assert layers and all(torch.count_nonzero(m.lora_B)==0 for m in layers.values())
    moments={};handles=[];mapping_checks=[]
    for name,m in layers.items():
        d=m.lora_A.shape[1]
        moments[name]={'activation':torch.zeros(d,d,device='cuda'),'delta':torch.zeros(d,d,device='cuda'),
                       'activation_sum':torch.zeros(d,device='cuda'),'delta_sum':torch.zeros(d,device='cuda'),'rows':0}
        def hook(layer,args,name=name):
            conv=layer.original_layer;x=args[0].detach().float();k=conv.kernel_size[0]
            assert conv.groups==1 and conv.kernel_size==(k,k)
            d=layer.lora_A.shape[1]
            patch=F.unfold(x,conv.kernel_size,dilation=conv.dilation,padding=conv.padding,stride=conv.stride).transpose(1,2)
            positions=torch.linspace(0,patch.shape[1]-1,min(256,patch.shape[1]),device=x.device).long()
            sampled=patch[:,positions,:].contiguous()
            # Official BiLaLoRA factors reshape W[out,in,k,k] as [out*k,in*k].
            # Split each im2col patch into exactly the same k contiguous chunks.
            rows=sampled.reshape(x.shape[0],-1,k,d)
            count=x.shape[0]//2
            a=rows[:count].reshape(-1,d);b=rows[count:].reshape(-1,d);delta=b-a
            state=moments[name]
            state['activation'].addmm_(a.T,a);state['delta'].addmm_(delta.T,delta)
            state['activation_sum']+=a.sum(0);state['delta_sum']+=delta.sum(0)
            state['rows']+=a.shape[0]
            if len(mapping_checks)<1:
                explicit=torch.einsum('bljd,ojd->blo',rows[:1,:16],conv.weight.reshape(conv.out_channels,k,d))
                if conv.bias is not None:explicit+=conv.bias
                actual=F.conv2d(x[:1],conv.weight,conv.bias,conv.stride,conv.padding,conv.dilation,conv.groups).flatten(2).transpose(1,2)[:,positions[:16]]
                error=(explicit-actual).abs().max().item()
                assert torch.allclose(explicit,actual,rtol=2e-4,atol=2e-4),(name,error)
                mapping_checks.append({'layer':name,'max_abs_error':error})
        handles.append(m.register_forward_pre_hook(hook))
    transform=lambda im:ToTensor()(CenterCrop(224)(Resize(256)(im)))
    normalize=Normalize((.48145466,.4578275,.40821073),(.26862954,.26130258,.27577711))
    dataset=train_subset.dataset
    selected=list(train_subset.indices[:32])
    was_training=model.training;model.eval();reference=None;probe=None
    for i in range(0,len(selected),4):
        images=[]
        for index in selected[i:i+4]:
            image=Image.open(Path(dataset.hazy_path)/dataset.hazy_image_list[index]).convert('RGB')
            images.append(transform(image))
        raw=torch.stack(images).cuda()
        # Fixed atmospheric scattering intervention, no clear GT or test data.
        rehe=raw*.7+.9*.3
        x=torch.cat([normalize(raw),normalize(rehe)],dim=0)
        output=model(x)
        if reference is None:reference=output.clone();probe=x.clone()
    for handle in handles:handle.remove()
    reports={}
    for name,layer in layers.items():
        stats=moments[name];rank,d=layer.lora_A.shape
        norms=layer.lora_A.norm(dim=1,keepdim=True).clone()
        extra={}
        if mode=='random_ortho':basis=torch.linalg.qr(layer.lora_A.float().T,mode='reduced').Q.T
        elif mode in ['eva_centered','delta_centered','gcpca']:
            def centered(key):
                mean=stats[key+'_sum'].double()/stats['rows']
                c=stats[key].double()/stats['rows']-mean[:,None]*mean[None,:]
                return (c+c.T)*.5
            ca,cb=centered('delta'),centered('activation')
            if mode in ['delta_centered','eva_centered']:
                values,vectors=torch.linalg.eigh(cb if mode=='eva_centered' else ca);basis=vectors[:,-rank:].flip(1).T.float().contiguous()
            else:
                # gcPCA v4: maximize v'(Ca-Cb)v / v'(Ca+Cb)v.
                # Drop numerical null dimensions, then keep the leading
                # generalized eigenspace. QR changes its basis, not its span.
                total=ca+cb;vals,vecs=torch.linalg.eigh((total+total.T)*.5)
                keep=vals>vals.max()*1e-6
                assert int(keep.sum())>=rank,(name,int(keep.sum()),rank)
                white=vecs[:,keep]*vals[keep].rsqrt()[None,:]
                contrast=white.T@(ca-cb)@white
                scores,axes=torch.linalg.eigh((contrast+contrast.T)*.5)
                directions=white@axes[:,-rank:].flip(1)
                basis=torch.linalg.qr(directions,mode='reduced').Q.T.float().contiguous()
                extra={'generalized_eigenvalue_min':scores[-rank].item(),'generalized_eigenvalue_max':scores[-1].item(),
                       'retained_denominator_dimensions':int(keep.sum()),'null_tolerance_relative':1e-6}
        else:
            cov=stats['activation' if mode=='eva' else 'delta']/stats['rows']
            values,vectors=torch.linalg.eigh((cov+cov.T)*.5)
            basis=vectors[:,-rank:].flip(1).T.contiguous()
        layer.lora_A.copy_(basis*norms)
        def fraction(key):
            c=stats[key];return ((basis@c)*basis).sum().div(c.trace().clamp_min(1e-20)).item()
        reports[name]={'A_shape':list(layer.lora_A.shape),'rows':stats['rows'],
                       'mean_row_norm':norms.mean().item(),'activation_energy_fraction':fraction('activation'),
                       'haze_response_energy_fraction':fraction('delta'),**extra}
    after=model(probe)
    invariant_error=(after-reference).abs().max().item()
    assert torch.allclose(after,reference,rtol=2e-5,atol=2e-5),invariant_error
    model.train(was_training)
    random.setstate(rng[0]);np.random.set_state(rng[1]);torch.random.set_rng_state(rng[2]);torch.cuda.set_rng_state_all(rng[3])
    report={'mode':mode,'images':[dataset.hazy_image_list[i] for i in selected],
            'calibration_scope':'first32 from training half only; center crop224 after short-side256',
            'rehaze':{'transmission':.7,'atmospheric_light':.9},'moment':('centered covariance' if mode in ['eva_centered','delta_centered','gcpca'] else 'uncentered second moment, EVA Theorem3.1'),
            'rank_redistribution':False,'normalization':'retain each original random A row norm',
            'mapping_checks':mapping_checks,'initial_output_max_abs_error':invariant_error,
            'seconds':time.perf_counter()-start,'layers':reports}
    Path('calibration_report.json').write_text(json.dumps(report,indent=2))
    print('EVA_CALIBRATION_COMPLETED',mode,report['seconds'],'seconds','initial_output_error',invariant_error,flush=True)
