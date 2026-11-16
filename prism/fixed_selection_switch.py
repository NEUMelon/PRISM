"""Control selection with the same canonical rebuild RNG advance in all arms."""
from pathlib import Path
import hashlib,json,random
import numpy as np
import torch

FIXED=['level3_block7','level3_block1','level3_block8']
def get_rng():
    return (random.getstate(),np.random.get_state(),torch.get_rng_state(),torch.cuda.get_rng_state_all())
def set_rng(s):
    random.setstate(s[0]);np.random.set_state(s[1]);torch.set_rng_state(s[2]);torch.cuda.set_rng_state_all(s[3])
def digest(s):
    return {'python':hashlib.sha256(repr(s[0]).encode()).hexdigest(),
            'numpy':hashlib.sha256(s[1][1].tobytes()+repr(s[1][2:]).encode()).hexdigest(),
            'torch_cpu':hashlib.sha256(s[2].numpy().tobytes()).hexdigest(),
            'torch_cuda':[hashlib.sha256(x.cpu().numpy().tobytes()).hexdigest() for x in s[3]]}

@torch.no_grad()
def controlled_switch(search_model,alphas,model_ctor,rebuild,extract,opt,base_checkpoint,strategy):
    assert strategy in ['dynamic','fixed']
    dynamic=sorted(alphas,key=lambda n:torch.sigmoid(alphas[n]).item(),reverse=True)[:opt.top_k_lora]
    old_state,_=extract(search_model,alphas,len(dynamic))
    fixed_state={n:p.detach().clone() for n,p in search_model.named_parameters()
                 if p.requires_grad and any(n.startswith(layer+'.') for layer in FIXED)}
    assert len(fixed_state)==15,len(fixed_state)
    source_norms={n:p.norm().item() for n,p in fixed_state.items() if n.endswith('lora_B')}
    assert all(v>0 for v in source_norms.values()),source_norms
    before=get_rng()
    def build(layers,state):
        base=model_ctor().to(opt.device)
        base.load_state_dict(torch.load(base_checkpoint,map_location=opt.device,weights_only=True))
        for p in base.parameters():p.requires_grad=False
        model,new_alphas=rebuild(base_model=base,selected_submodules=layers,lora_state_dict=state,r=opt.lora_r,lora_alpha=opt.lora_alpha)
        return model.to(opt.device),new_alphas
    dynamic_model,dynamic_alphas=build(dynamic,old_state)
    after_dynamic=get_rng()
    # The dynamic selection can differ across initialization methods. Using its
    # RNG advance would silently give the two initializers different crop streams.
    # Always advance as the same published fixed architecture in every arm.
    set_rng(before)
    fixed_model,fixed_alphas=build(FIXED,fixed_state)
    after_canonical=get_rng()
    current=dict(fixed_model.named_parameters())
    assert all(n in current and torch.equal(current[n],p) for n,p in fixed_state.items()),'Search-trained adapter state was not preserved'
    if strategy=='dynamic':
        model,new_alphas,selected=dynamic_model,dynamic_alphas,dynamic
        del fixed_model,fixed_alphas
    else:
        model,new_alphas=fixed_model,fixed_alphas
        selected=FIXED.copy();del dynamic_model,dynamic_alphas
    set_rng(after_canonical)
    assert digest(get_rng())==digest(after_canonical)
    params=sum(p.numel() for p in model.parameters() if p.requires_grad)
    if strategy=='fixed':assert params==110595,params
    audit={'strategy':strategy,'dynamic_selected':dynamic,'selected':selected,'trainable_parameters':params,
           'search_trained_state_preserved':True,'source_fixed_B_norms':source_norms,
           'rng_before':digest(before),'rng_after_dynamic':digest(after_dynamic),'rng_after_canonical':digest(after_canonical),
           'rng_final':digest(get_rng()),'fixed_uses_published_layers':strategy=='fixed',
           'rebuild_rng_policy':'published_fixed_layers_for_all_arms_v2'}
    Path('switch_audit.json').write_text(json.dumps(audit,indent=2))
    print('CONTROLLED_SWITCH',strategy,selected,params,flush=True)
    return model,new_alphas,selected

def trace_batch(epoch,batch,x,split='train'):
    # Hash already-augmented input tensors to directly check paired sample flow.
    digest_x=hashlib.sha256(x.detach().contiguous().cpu().numpy().tobytes()).hexdigest()
    with Path('batch_trace.jsonl').open('a') as f:
        f.write(json.dumps({'epoch':epoch,'batch':batch,'split':split,'sha256':digest_x})+'\n')
