"""Render every deliverable page and verify the manuscript's data and formatting."""
from pathlib import Path
import hashlib,json,re
import fitz
import numpy as np
from PIL import Image,ImageDraw

P=Path(__file__).resolve().parent
FORBIDDEN=['本轮','老师','网页报告','归档','第2版','第 2 版','URHI150','尚未取得','仍待补充','QALA']

def main():
    preview_dir = P/'preview'
    preview_dir.mkdir(exist_ok=True)
    for old in preview_dir.glob('*.png'):
        old.unlink()
    report={'gpu_used':False,'pdfs':{},'data':{}}
    for tag,stem in [('main','PRISM_中文论文_20261116'),('supplement','PRISM_补充材料_20261116')]:
        pdf=P/(stem+'.pdf');log=(P/(stem+'.log')).read_text(encoding='utf-8',errors='replace')
        assert not re.search(r'Overfull \\[hv]box|undefined references|Citation .+ undefined|Reference .+ undefined|multiply defined',log), (tag, 'LaTeX log warning/error')
        with fitz.open(pdf) as doc:
            checks=[];thumbs=[];texts=[]
            for i,page in enumerate(doc):
                text=page.get_text();texts.append(text)
                assert '\ufffd' not in text
                bad=[list(b[:4]) for b in page.get_text('blocks') if b[0]<-.5 or b[1]<-.5 or b[2]>page.rect.width+.5 or b[3]>page.rect.height+.5]
                assert not bad,(tag,i,bad)
                checks.append({'page':i+1,'characters':len(text),'out_of_page':bad})
                pix=page.get_pixmap(matrix=fitz.Matrix(1.8,1.8));pix.save(preview_dir/f'{tag}-{i+1}.png')
                im=Image.frombytes('RGB',[pix.width,pix.height],pix.samples);im.thumbnail((440,590));canvas=Image.new('RGB',(460,620),'#d9dce2');canvas.paste(im,((460-im.width)//2,12));ImageDraw.Draw(canvas).text((12,600),f'{tag} {i+1}',fill='black');thumbs.append(canvas)
            fulltext='\n'.join(texts)
            for word in FORBIDDEN:assert word not in fulltext,(tag,word)
            assert 'References' not in fulltext
            contact=Image.new('RGB',(1380,620*((len(doc)+2)//3)),'#d9dce2')
            for i,im in enumerate(thumbs):contact.paste(im,((i%3)*460,(i//3)*620))
            contact.save(preview_dir/(tag+'-contact.png'))
            (preview_dir/(tag+'.txt')).write_text(fulltext,encoding='utf-8')
            refs=len(re.findall(r'\\bibitem\[',(P/(stem+'.bbl')).read_text(encoding='utf-8')))
            out={'path':pdf.name,'pages':len(doc),'references':refs,'bytes':pdf.stat().st_size,'sha256':hashlib.sha256(pdf.read_bytes()).hexdigest(),'page_checks':checks,'overfull_boxes':False,'undefined_references':False}
            if tag=='main':
                assert '摘要' in texts[0] and '关键词：' in texts[0] and 'Abstract' not in texts[0]
                assert '——面向真实图像去雾的评训分离适配' in texts[0]
                spans=[s for b in doc[0].get_text('dict')['blocks'] for l in b.get('lines',[]) for s in l['spans'] if s['bbox'][1]<140]
                out['title_spans']=[{'text':s['text'],'size':s['size']} for s in spans]
                assert max(s['size'] for s in spans)-min(s['size'] for s in spans)<.02
                labels=[s for b in doc[1].get_text('dict')['blocks'] for l in b.get('lines',[]) for s in l['spans'] if s['bbox'][1]<370 and re.search('[\u4e00-\u9fff]',s['text'])]
                out['method_chinese_label_min_pt']=min(s['size'] for s in labels)
                assert out['method_chinese_label_min_pt']>=7.5
                table_page=next(i for i,t in enumerate(texts) if 'RIDCP' in t and 'PHATNet' in t)
                im=Image.open(preview_dir/f'main-{table_page+1}.png').convert('L');im.save(preview_dir/'main-table-grayscale.png')
                doc[1].get_pixmap(matrix=fitz.Matrix(3,3),clip=fitz.Rect(55,69,558,370)).save(preview_dir/'method.png')
                # Render Fig 2 region for quick inspection
                landscape_page=next(i for i,t in enumerate(texts) if 'Perceptually Feasible Region' in t)
                doc[landscape_page].get_pixmap(matrix=fitz.Matrix(3,3),clip=fitz.Rect(45,60,570,295)).save(preview_dir/'fig2_landscape.png')
            report['pdfs'][tag]=out

    sd_count=0
    for f in (P/'generated').glob('*.tex'):
        for mean,sd in re.findall(r'\\sd\{([^}]+)\}\{([^}]+)\}',f.read_text(encoding='utf-8')):
            assert len(mean.partition('.')[2])==len(sd.partition('.')[2]),(f.name,mean,sd)
            sd_count+=1

    # Verify main_table.tex rankings
    main_table=(P/'generated/main_table.tex').read_text(encoding='utf-8')
    rows=[line.split(' & ') for line in main_table.splitlines() if ' & ' in line and len(line.split(' & '))==7 and not line.startswith(' & ')]
    assert len(rows)==9
    vals=[]
    for row in rows:
        vals.append([float(re.search(r'[-+]?\d+\.\d+',cell).group()) for cell in row[1:]])
    for j in range(6):
        ranked=sorted({v[j] for v in vals},reverse=j%3!=0)
        for row,v in zip(rows,vals):
            assert ('\\best{' in row[j+1])==(v[j]==ranked[0])
            assert ('\\runnerup{' in row[j+1])==(v[j]==ranked[1])

    # Verify course10_table.tex rankings
    c10_table=(P/'generated/course10_table.tex').read_text(encoding='utf-8')
    c10_rows=[line.split(' & ') for line in c10_table.splitlines() if ' & ' in line and len(line.split(' & '))==4 and not line.startswith('方法')]
    assert len(c10_rows)==4
    c10_vals=[[float(re.search(r'[-+]?\d+\.\d+',cell).group()) for cell in row[1:]] for row in c10_rows]
    for j in range(3):
        ranked=sorted({v[j] for v in c10_vals},reverse=j!=0)
        for row,v in zip(c10_rows,c10_vals):
            assert ('\\best{' in row[j+1])==(v[j]==ranked[0])
            assert ('\\runnerup{' in row[j+1])==(v[j]==ranked[1])

    prov=json.loads((P/'evidence/figure_provenance.json').read_text(encoding='utf-8'))
    for x in prov['images']:
        src=P.parent/x['source'];assert hashlib.sha256(src.read_bytes()).hexdigest()==x['source_sha256']
        with Image.open(P/'figures'/f'{x["case"]}_{x["model"]}_crop.png') as im:assert im.width==2*im.height
    for case in ['swan','snow','buildings','train']:
        b=np.asarray(Image.open(P/'figures'/f'{case}_warm_published_3407_crop.png').convert('RGB'),dtype=np.float32)
        o=np.asarray(Image.open(P/'figures'/f'{case}_warm_quality_constrained_3407_crop.png').convert('RGB'),dtype=np.float32)
        expected=np.clip(8*np.mean(np.abs(o-b),axis=2),0,255).astype(np.uint8)
        assert np.array_equal(expected,np.asarray(Image.open(P/'figures'/f'{case}_difference.png')))

    audit=json.loads((P/'evidence/generation_audit.json').read_text())
    assert audit['abstract_characters']<350 and audit['source_rows']==30471
    for f in ['main.tex','supplement.tex','preamble.tex']:
        text=(P/f).read_text(encoding='utf-8')
        assert not re.search(r'(?<![~\\])\\cite\{',text)
        assert 'QALA' not in text, f
    assert not re.search(r'H\(S\)|Q\(S\)|V_q|\$s=2\$|QALA',(P/'figures/method_architecture.tex').read_text(encoding='utf-8'))
    report['data']=dict(audit,sd_cells_checked=sd_count,mean_sd_precision_consistent=True,rank_cells_checked=54+12,real_image_sources_checked=len(prov['images']),crop_aspect_ratio='2:1',difference_maps_verified=4,source_results_sha256=hashlib.sha256((P.parent/'main_table_round/analysis.json').read_bytes()).hexdigest())
    (P/'evidence/final_checks.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:{'pages':v['pages'],'references':v['references']} for k,v in report['pdfs'].items()},ensure_ascii=False))
    print('Checked',sd_count,'mean/SD cells, 66 ranking cells and',len(prov['images']),'original image sources.')

if __name__=='__main__':main()
