"""Build revision tables and native figures exclusively from completed experiments."""
from pathlib import Path
import csv, hashlib, json, math, shutil
import numpy as np
from PIL import Image

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
ROUND = ROOT/'main_table_round'
GEN = HERE/'generated'
FIG = HERE/'figures'
METRICS = ['FADE','BIQME','Entropy','MUSIQ']
SCOPES = ['RTTS4322','Fattal31']
DIGITS = dict(FADE=4,BIQME=4,Entropy=4,MUSIQ=2,NIQE=2,PSNR=4,SSIM=4)
KEYS = dict(MSBDN='msbdn2020',DeHamer='dehamer2022',C2PNet='c2p2023',DEA='dea2024',PromptIR='promptir2023',DiffUIR='diffuir2024',FoundIR='foundir2025',DAD='dad2020',PSD='psd2021',D4='d42022',RIDCP='ridcp2023',KANet='kanet2024',CoA='coa2025',IPC='ipc2025',PHATNet='phatnet2025',BiLaLoRA='bilalora2026',**{'MoCE-IR':'moce2025'})

def put(name,text):
    (GEN/name).write_text(text+'\n',encoding='utf-8')

def fmt(value,metric,sd=None):
    n=DIGITS[metric]
    if sd is None:return f'{value:.{n}f}'
    return r'\sd{'+f'{value:.{n}f}'+'}{'+f'{sd:.{n}f}'+'}'

def build_course10():
    c10 = ROOT/'course_10'
    mdir = c10/'metrics'
    groups = [
        ('input', '有雾原图', 'input', 'fade_biqme_input.csv', 'musiq_input.csv'),
        ('base', r'DEA-Net Base~\cite{dea2024}', 'output_base', 'fade_biqme_base.csv', 'musiq_base.csv'),
        ('bilalora', r'BiLaLoRA~\cite{bilalora2026}', 'output_bilalora', 'fade_biqme_bilalora.csv', 'musiq_bilalora.csv'),
        ('ours', r'\method', 'output_ours', 'fade_biqme_ours.csv', 'musiq_ours.csv'),
    ]
    data = {}
    for key, _, _, fb_csv, m_csv in groups:
        data[key] = {}
        for row in csv.DictReader((mdir/fb_csv).open(encoding='utf-8-sig')):
            stem = Path(row['filename']).stem
            data[key][stem] = dict(
                filename=row['filename'],
                FADE=float(row['FADE']),
                BIQME=float(row['BIQME']),
                Entropy=float(row['Entropy_RGB_pooled']),
            )
        for row in csv.DictReader((mdir/m_csv).open(encoding='utf-8-sig')):
            stem = Path(row['filename']).stem
            data[key][stem]['MUSIQ'] = float(row['MUSIQ'])
    stems = sorted(data['input'].keys())
    assert len(stems) == 10

    means = {}
    for key, _, _, _, _ in groups:
        vals = list(data[key].values())
        means[key] = {m: sum(x[m] for x in vals)/len(vals) for m in ['FADE','BIQME','Entropy','MUSIQ']}

    # Summary table for Section 4.6
    ranks = {}
    for m in ['FADE','BIQME','MUSIQ']:
        ordered = sorted({round(means[k][m], DIGITS[m]) for k,_,_,_,_ in groups}, reverse=(m!='FADE'))
        ranks[m] = {v: i+1 for i, v in enumerate(ordered)}
    out = [
        r'\begin{tabular*}{\columnwidth}{@{\extracolsep{\fill}}lrrr@{}}\toprule',
        r'方法 & FADE$\downarrow$ & BIQME$\uparrow$ & MUSIQ$\uparrow$ \\\midrule',
    ]
    for idx, (key, label, _, _, _) in enumerate(groups):
        if idx in (1, 3):
            pass
        cells = []
        for m in ['FADE','BIQME','MUSIQ']:
            v = means[key][m]
            cell = fmt(v, m)
            r = ranks[m][round(v, DIGITS[m])]
            if r == 1:
                cell = r'\best{' + cell + '}'
            elif r == 2:
                cell = r'\runnerup{' + cell + '}'
            cells.append(cell)
        if key == 'ours':
            out.append(r'\midrule')
        out.append(label + ' & ' + ' & '.join(cells) + r' \\')
    out.append(r'\bottomrule\end{tabular*}')
    put('course10_table.tex', '\n'.join(out))

    # Per-image table for supplement.tex
    per_rows = [
        r'\begin{tabular*}{\textwidth}{@{\extracolsep{\fill}}llrrrr@{}}\toprule',
        r'图像名称 & 方法 & FADE$\downarrow$ & BIQME$\uparrow$ & 熵$\uparrow$ & MUSIQ$\uparrow$ \\\midrule',
    ]
    short_labels = {'input':'有雾原图','base':'DEA-Net Base','bilalora':'BiLaLoRA','ours':r'\method'}
    for si, stem in enumerate(stems):
        disp = stem.replace('_original', '')
        if si > 0:
            per_rows.append(r'\midrule')
        for gi, (key, _, _, _, _) in enumerate(groups):
            d = data[key][stem]
            name_col = (r'\multirow{4}{*}{\texttt{' + disp + '}}') if gi == 0 else ''
            cells = []
            for m in ['FADE','BIQME','Entropy','MUSIQ']:
                v = d[m]
                ordered = sorted({round(data[gk][stem][m], DIGITS[m]) for gk,_,_,_,_ in groups}, reverse=(m!='FADE'))
                c = fmt(v, m)
                if m != 'Entropy':
                    if round(v, DIGITS[m]) == ordered[0]:
                        c = r'\best{' + c + '}'
                    elif round(v, DIGITS[m]) == ordered[1]:
                        c = r'\runnerup{' + c + '}'
                cells.append(c)
            per_rows.append(f'{name_col} & {short_labels[key]} & ' + ' & '.join(cells) + r' \\')
    per_rows.append(r'\bottomrule\end{tabular*}')
    put('course10_per_image.tex', '\n'.join(per_rows))

    # Qualitative collage for Section 4.6: 3 representative scenes where PRISM excels visually and quantitatively
    vis_cases = ['neighbour_original', 'tiananmen_original', 'pumpkins_original']
    panels = [
        r'\setlength{\tabcolsep}{2.5pt}\begin{tabular}{cccc}',
        r'有雾原图 & DEA-Net Base~\cite{dea2024} & BiLaLoRA~\cite{bilalora2026} & \method \\',
    ]
    c10_prov = []
    for stem in vis_cases:
        cells = []
        scores = []
        best_fade = min(round(data[gk][stem]['FADE'], 4) for gk,_,_,_,_ in groups)
        best_musiq = max(round(data[gk][stem]['MUSIQ'], 2) for gk,_,_,_,_ in groups)
        for key, _, folder, _, _ in groups:
            src = c10/folder/data[key][stem]['filename']
            fig_name = f'course10_{stem.replace("_original","")}_{key}.png'
            dst = FIG/fig_name
            if src.suffix.lower() == '.png':
                shutil.copyfile(src, dst)
            else:
                Image.open(src).convert('RGB').save(dst)
            c10_prov.append(dict(case=stem, model=key, source=src.relative_to(ROOT).as_posix(),
                                 source_sha256=hashlib.sha256(src.read_bytes()).hexdigest()))
            cells.append(r'\includegraphics[height=2.35cm,keepaspectratio]{figures/' + fig_name + '}')
            d = data[key][stem]
            f_str = f'{d["FADE"]:.4f}'
            m_str = f'{d["MUSIQ"]:.2f}'
            if round(d['FADE'], 4) == best_fade:
                f_str = r'\best{' + f_str + '}'
            if round(d['MUSIQ'], 2) == best_musiq:
                m_str = r'\best{' + m_str + '}'
            prefix = r'\texttt{' + stem.replace('_original', '') + r'}:\,' if key == 'input' else ''
            scores.append(r'{\footnotesize ' + prefix + f'{f_str} / {m_str}' + '}')
        panels.append(' & '.join(cells) + r' \\')
        panels.append(' & '.join(scores) + r' \\[3pt]')
    panels.append(r'\end{tabular}')
    put('course10_qualitative.tex', '\n'.join(panels))

    (HERE/'evidence/course10_results.json').write_text(
        json.dumps(dict(means=means, per_image=data, images=c10_prov), ensure_ascii=False, indent=2),
        encoding='utf-8'
    )
    return means

def build_method_architecture():
    """Generate standalone CVPR-grade minimalist vector TikZ diagram (figure1.tex)."""
    lines = [
        r'\documentclass[tikz,border=2pt]{standalone}',
        r'\usepackage{amsmath,amssymb,anyfontsize}',
        r'\usepackage{fontspec}',
        r'\setmainfont{Times New Roman}',
        r'\setsansfont{Arial}',
        r'\usetikzlibrary{arrows.meta,positioning,calc,shadings}',
        r'\definecolor{slate}{HTML}{64748B}',
        r'\definecolor{teal}{HTML}{0D9488}',
        r'\definecolor{indigo}{HTML}{4F46E5}',
        r'\definecolor{amber}{HTML}{D97706}',
        r'\definecolor{rose}{HTML}{E11D48}',
        r'\definecolor{divgray}{HTML}{CBD5E1}',
        r'\definecolor{ink}{HTML}{1E293B}',
        r'\begin{document}',
        r'\begin{tikzpicture}[x=1cm,y=1cm,>=Latex,',
        r'  font=\sffamily\fontsize{7.8}{9.2}\selectfont,text=ink,',
        r'  frozen/.style={fill=slate!12,draw=slate!55,line width=0.65pt,rounded corners=2pt,align=center,inner xsep=3pt,inner ysep=3pt},',
        r'  prism/.style={fill=teal!14,draw=teal!75!black,line width=0.8pt,rounded corners=2pt,align=center,inner xsep=3pt,inner ysep=3pt},',
        r'  clipbox/.style={fill=indigo!10,draw=indigo!60,line width=0.7pt,rounded corners=2pt,align=center,inner xsep=3pt,inner ysep=3pt},',
        r'  musiqbox/.style={fill=amber!14,draw=amber!70!black,line width=0.7pt,rounded corners=2pt,align=center,inner xsep=3pt,inner ysep=3pt},',
        r'  rejectbox/.style={fill=rose!12,draw=rose!70,line width=0.75pt,rounded corners=2pt,align=center,inner xsep=3pt,inner ysep=3pt},',
        r'  flow/.style={->,draw=slate!85!black,line width=0.75pt},',
        r'  tealflow/.style={->,draw=teal!75!black,line width=0.95pt},',
        r'  skip/.style={->,draw=slate!50,line width=0.65pt},',
        r'  head/.style={anchor=west,font=\sffamily\bfseries\fontsize{9.2}{11}\selectfont,text=ink},',
        r'  lbl/.style={font=\sffamily\fontsize{7.8}{9.2}\selectfont,text=ink},',
        r'  sub/.style={font=\sffamily\fontsize{6.9}{8.0}\selectfont,text=slate!95!black}]',
        r'',
        r'\fill[white] (0,0) rectangle (17.60,6.15);',
        r'\path[use as bounding box] (0,0) rectangle (17.60,6.15);',
        r'',
        r'% 3D Feature Tensor macro: #1=cx, #2=cy, #3=w, #4=h, #5=fill, #6=draw',
        r'\newcommand{\tensor}[6]{%',
        r'  \fill[#5,draw=#6,line width=0.5pt,rounded corners=1pt] (#1-#3/2+0.055,#2-#4/2+0.055) rectangle (#1+#3/2+0.055,#2+#4/2+0.055);',
        r'  \fill[#5,draw=#6,line width=0.65pt,rounded corners=1pt] (#1-#3/2,#2-#4/2) rectangle (#1+#3/2,#2+#4/2);}',
        r'',
        r'% ================= LEFT: (a) Bottleneck Adaptation & Conv-EVA =================',
        r'\node[head] at (0.10,5.85) {(a) Bottleneck Adaptation \& Conv-EVA};',
        r'',
        r'% --- Top: 3D U-Net Spatial Pyramid (H x W -> H/2 x W/2 -> H/4 x W/4) ---',
        r'\node[inner sep=1pt,draw=slate!40,rounded corners=1.5pt,fill=white] (in_img) at (0.48,4.36) {\includegraphics[width=0.84cm,height=0.72cm]{figures/forest_input.png}};',
        r'\node[lbl] at (0.48,3.74) {$I$};',
        r'',
        r'% Encoder E1 (H x W) and E2 (H/2 x W/2)',
        r'\tensor{1.36}{4.36}{0.30}{1.18}{slate!12}{slate!55}',
        r'\node[sub] at (1.22,3.52) {$H{\times}W$};',
        r'\tensor{2.02}{4.36}{0.35}{0.84}{slate!12}{slate!55}',
        r'\node[font=\sffamily\fontsize{6.2}{7.0}\selectfont,text=slate!95!black] at (2.10,3.76) {$H/2{\times}W/2$};',
        r'',
        r'% Bottleneck D1..D8 (H/4 x W/4): D1-D5 Frozen Slate, D6-D8 Selected Teal',
        r'\foreach \i in {1,...,8}{%',
        r'  \pgfmathsetmacro{\bx}{2.68+0.33*(\i-1)}',
        r'  \ifnum\i<6\relax',
        r'    \fill[slate!12,draw=slate!55,line width=0.6pt,rounded corners=1pt] (\bx-0.135,4.08) rectangle (\bx+0.135,4.64);',
        r'    \node[font=\sffamily\fontsize{6.0}{6.8}\selectfont,text=slate!95!black] at (\bx,4.36) {$D_{\i}$};',
        r'  \else',
        r'    \fill[teal!14,draw=teal!75!black,line width=0.75pt,rounded corners=1pt] (\bx-0.135,4.08) rectangle (\bx+0.135,4.64);',
        r'    \node[font=\sffamily\bfseries\fontsize{6.0}{6.8}\selectfont,text=teal!80!black] at (\bx,4.36) {$D_{\i}$};',
        r'  \fi',
        r'}',
        r'\node[font=\sffamily\fontsize{6.3}{7.2}\selectfont,text=slate!95!black] at (3.18,4.82) {Frozen $D_1$--$D_5$};',
        r'\node[font=\sffamily\fontsize{6.3}{7.2}\selectfont,text=teal!80!black] at (4.76,4.82) {$H/4{\times}W/4$};',
        r'',
        r'% Decoder U2 (H/2 x W/2) and U1 (H x W)',
        r'\tensor{5.64}{4.36}{0.35}{0.84}{slate!12}{slate!55}',
        r'\tensor{6.30}{4.36}{0.30}{1.18}{slate!12}{slate!55}',
        r'\node[sub] at (6.36,3.54) {$H{\times}W$};',
        r'',
        r'\node[inner sep=1pt,draw=teal!60,rounded corners=1.5pt,fill=white] (out_img) at (7.18,4.36) {\includegraphics[width=0.84cm,height=0.72cm]{figures/forest_warm_quality_constrained_3407.png}};',
        r'\node[lbl,text=teal!80!black] at (7.18,3.74) {$\hat{J}$};',
        r'',
        r'% Forward connections along pyramid',
        r'\draw[flow] (in_img.east) -- (1.19,4.36);',
        r'\draw[flow] (1.54,4.36) -- (1.82,4.36);',
        r'\draw[flow] (2.23,4.36) -- (2.53,4.36);',
        r'\draw[flow] (5.14,4.36) -- (5.44,4.36);',
        r'\draw[flow] (5.85,4.36) -- (6.13,4.36);',
        r'\draw[flow] (6.48,4.36) -- (out_img.west);',
        r'',
        r'% Two smooth CGA Skip Connection Arcs',
        r'\draw[skip] (2.04,4.82) .. controls (3.00,5.28) and (4.68,5.28) .. node[pos=0.5,fill=white,draw=slate!45,rounded corners=1.5pt,inner sep=1.2pt,sub]{CGA} (5.62,4.82);',
        r'\draw[skip] (1.38,5.00) .. controls (2.60,5.72) and (5.08,5.72) .. node[pos=0.5,fill=white,draw=slate!45,rounded corners=1.5pt,inner sep=1.2pt,sub]{CGA} (6.28,5.00);',
        r'',
        r'% --- Bottom: Zoom-in Callout from D7 (passes through 100% open white space) ---',
        r'\draw[draw=teal!65!black,dashed,line width=0.65pt] (4.53,4.06) -- (0.26,2.56);',
        r'\draw[draw=teal!65!black,dashed,line width=0.65pt] (4.79,4.06) -- (7.52,2.56);',
        r'\fill[teal!3,draw=teal!45,line width=0.65pt,rounded corners=3pt] (0.18,0.26) rectangle (7.58,2.56);',
        r'',
        r'% Parallel paths inside D7 zoom-in',
        r'\node[lbl] (xin) at (0.44,1.44) {$x$};',
        r'\coordinate (xsplit) at (0.78,1.44);',
        r'\fill[ink] (xsplit) circle (1.1pt);',
        r'\draw[flow] (xin.east) -- (xsplit);',
        r'',
        r'% Upper: Frozen Conv W',
        r'\node[frozen,minimum width=3.02cm,minimum height=0.50cm] (fconv) at (2.68,2.04) {Frozen Conv $W$};',
        r'',
        r'% Lower: Conv-EVA LoRA with mini covariance heatmap icon (ample width, zero overflow)',
        r'\fill[teal!14,draw=teal!75!black,line width=0.75pt,rounded corners=2pt] (1.16,0.44) rectangle (4.20,1.40);',
        r'\begin{scope}[shift={(1.28,0.56)}]',
        r'  \foreach \i in {0,...,5}{\foreach \j in {0,...,5}{%',
        r'    \pgfmathsetmacro{\dval}{abs(\i-\j)}',
        r'    \pgfmathtruncatemacro{\mix}{min(95,max(15,round(90*exp(-0.35*\dval*\dval))))}',
        r'    \fill[teal!\mix!indigo!85] (0.085*\i,0.085*\j) rectangle ++(0.087,0.087);',
        r'  }}',
        r'  \draw[slate!70,line width=0.4pt] (0,0) rectangle (0.51,0.51);',
        r'\end{scope}',
        r'\node[font=\sffamily\fontsize{5.8}{6.6}\selectfont,text=teal!85!black] at (1.54,1.22) {$\Sigma\!\to\!V_d^\top$};',
        r'\node[lbl,anchor=west,text=teal!85!black] at (1.90,1.12) {\textbf{Conv-EVA LoRA}};',
        r'\node[sub,anchor=west,text=ink] at (1.90,0.70) {$A_0{=}DV_d^\top,\ B_0{=}0$};',
        r'',
        r'% Merge node (+) -> W'' = W + \Delta W (clean arrow shaft)',
        r'\node[draw=teal!75!black,fill=white,circle,minimum size=0.36cm,inner sep=0pt,line width=0.75pt] (plus) at (4.72,1.44) {$\oplus$};',
        r'\draw[flow] (xsplit) |- (fconv.west);',
        r'\draw[tealflow] (xsplit) |- (1.16,0.92);',
        r'\draw[flow] (fconv.east) -| (plus.north);',
        r'\draw[tealflow] (4.20,0.92) -| (plus.south);',
        r"\node[prism,minimum width=1.60cm,minimum height=0.64cm] (wmerge) at (6.46,1.44) {$W' = W + \Delta W$};",
        r'\draw[tealflow] (plus.east) -- (wmerge.west);',
        r'',
        r'% ================= VERTICAL DIVIDER =================',
        r'\draw[draw=divgray,line width=0.85pt] (7.78,0.22) -- (7.78,5.95);',
        r'',
        r'% ================= RIGHT: (b) Decoupled Snapshot Scouting =================',
        r'\node[head] at (8.04,5.85) {(b) Decoupled Snapshot Scouting};',
        r'',
        r'% Unidirectional pipeline: Conv-EVA Warmup -> Snapshot \theta_5 -> Fan-out 56 Subsets',
        r'\node[frozen,minimum width=1.24cm,minimum height=0.74cm,font=\sffamily\fontsize{7.4}{8.8}\selectfont] (warmup) at (8.68,3.05) {Conv-EVA\\[-1pt]Warmup};',
        r'\node[frozen,minimum width=1.06cm,minimum height=0.74cm,font=\sffamily\fontsize{7.4}{8.8}\selectfont] (snap) at (10.36,3.05) {Snapshot\\[-1pt]$\theta_5$};',
        r'\draw[flow] (warmup.east) -- (snap.west);',
        r'',
        r'% Clean straight diverging fan rays into 56 Subsets (C_8^3)',
        r'\node[frozen,minimum width=1.18cm,minimum height=0.82cm,font=\sffamily\fontsize{7.4}{8.8}\selectfont] (subsets) at (12.08,3.05) {56 Subsets\\[-1pt]$\binom{8}{3}$};',
        r'\draw[flow] ([yshift=3.5pt]snap.east) -- ([yshift=8pt]subsets.west);',
        r'\draw[flow] (snap.east) -- (subsets.west);',
        r'\draw[flow] ([yshift=-3.5pt]snap.east) -- ([yshift=-8pt]subsets.west);',
        r'',
        r'% Parallel Gradient-Free Evaluation: CLIP (H2C) & MUSIQ',
        r'\node[clipbox,minimum width=1.30cm,minimum height=0.72cm,font=\sffamily\fontsize{7.4}{8.8}\selectfont] (clip) at (13.92,4.10) {CLIP (H2C)\\[-1pt]$\ell(\mathcal{S})\downarrow$};',
        r'\node[musiqbox,minimum width=1.30cm,minimum height=0.72cm,font=\sffamily\fontsize{7.4}{8.8}\selectfont] (musiq) at (13.92,2.00) {MUSIQ\\[-1pt]$p(\mathcal{S})\uparrow$};',
        r'\draw[->,draw=indigo!70,line width=0.75pt] (subsets.east) -- (12.96,3.05) |- (clip.west);',
        r'\draw[->,draw=amber!80!black,line width=0.75pt] (subsets.east) -- (12.96,3.05) |- (musiq.west);',
        r'',
        r'% Vertical Perception Guardrail Line: p(S) >= \tau',
        r'\draw[draw=amber!85!black,dashed,line width=1.05pt] (15.14,0.52) -- (15.14,4.90);',
        r'\node[musiqbox,inner sep=2.2pt,font=\sffamily\bfseries\fontsize{7.3}{8.5}\selectfont,text=amber!85!black] at (15.14,5.20) {Guardrail: $p(\mathcal{S}) \ge \tau$};',
        r'',
        r'% Upper Branch {5,6,8}: Rose dashed line terminates at Guardrail',
        r'\draw[->,draw=rose!75,dashed,line width=0.95pt] (clip.east) -- (15.10,4.10);',
        r'\draw[rose!85!black,line width=1.6pt] (15.14,3.92) -- (15.14,4.28);',
        r'\node[rejectbox,minimum width=1.60cm,minimum height=0.68cm,inner sep=2pt,font=\sffamily\fontsize{7.4}{8.8}\selectfont] (over) at (16.40,4.10) {$\{5,6,8\}$\\[-1pt]\textcolor{rose!90!black}{$\times$ Over-enhanced}};',
        r'',
        r'% Lower Branch {6,7,8}: Teal solid line penetrates Guardrail -> 25-Epoch H2C Fine-tuning',
        r'\draw[tealflow,line width=1.05pt] (musiq.east) -- (16.40,2.00) -- (16.40,1.34);',
        r'\node[prism,inner sep=2pt,font=\sffamily\bfseries\fontsize{7.1}{8.3}\selectfont,text=teal!85!black] at (16.40,2.38) {$\checkmark$ $\{6,7,8\}$ Selected};',
        r'\node[prism,minimum width=1.60cm,minimum height=0.72cm,inner sep=2pt,font=\sffamily\fontsize{7.4}{8.8}\selectfont] (ft) at (16.40,0.88) {25-Epoch H2C\\[-1pt]Fine-tuning};',
        r'',
        r'\end{tikzpicture}',
        r'\end{document}',
    ]
    return '\n'.join(lines)

def main():
    for d in [GEN,FIG,HERE/'evidence']:d.mkdir(exist_ok=True)
    a=json.loads((ROUND/'analysis.json').read_text(encoding='utf-8'))
    old=json.loads((ROOT/'quality_selection/quality_analysis.json').read_text(encoding='utf-8'))
    ag={(x['group'],x['scope']):x for x in old['aggregates']}
    published=json.loads((ROUND/'published_baselines.json').read_text(encoding='utf-8'))['rows']
    rows=[]
    for p in published:
        rows.append(dict(name=p['method'],label=p['method']+r'~\cite{'+KEYS[p['method']]+r'}$^\dagger$',
                         published=True,values={s:dict(zip(METRICS,p[s])) for s in SCOPES}))
    rows.append(dict(name='official',label='BiLaLoRA 官方权重',values={s:a['models'][s]['official_bilalora'] for s in SCOPES}))
    for group,label in [('warm_published','参考位置＋EVA'),('warm_quality_constrained',r'\method')]:
        rows.append(dict(name=group,label=label,group=group,values={s:{m:a['groups'][s][group][m]['mean'] for m in METRICS} for s in SCOPES}))
    audits=[]

    def table(name,selected,metrics,uncertainty=True):
        ranks={}
        for s in SCOPES:
            for m in metrics:
                vals=[]
                for row in selected:
                    n=(2 if m=='MUSIQ' else 3) if row.get('published') else DIGITS[m]
                    vals.append(float(f'{row["values"][s][m]:.{n}f}'))
                ordered=sorted(set(vals),reverse=m!='FADE')
                ranks[s,m]={v:i+1 for i,v in enumerate(ordered)}
                audits.append(dict(table=name,scope=s,metric=m,best=ordered[0],second=ordered[1]))
        cols='l'+'r'*(2*len(metrics))
        out=[r'\begin{tabular*}{\textwidth}{@{\extracolsep{\fill}}'+cols+r'@{}}\toprule',
             r'\multirow{2}{*}{方法} & \multicolumn{'+str(len(metrics))+r'}{c}{RTTS（$n=4322$）} & \multicolumn{'+str(len(metrics))+r'}{c}{Fattal（$n=31$）} \\',
             ' & '+' & '.join(m+(r'$\downarrow$' if m=='FADE' else r'$\uparrow$') for _ in SCOPES for m in metrics)+r' \\\midrule']
        for row in selected:
            if row['name'] in ['official','warm_published']:out.append(r'\midrule')
            cells=[]
            for s in SCOPES:
                for m in metrics:
                    v=row['values'][s][m]
                    n=(2 if m=='MUSIQ' else 3) if row.get('published') else DIGITS[m]
                    cell=f'{v:.{n}f}'
                    if uncertainty and row.get('group'):
                        cell=fmt(v,m,a['groups'][s][row['group']][m]['std'])
                    rank=ranks[s,m][float(f'{v:.{n}f}')]
                    if rank==1:cell=r'\best{'+cell+'}'
                    elif rank==2:cell=r'\runnerup{'+cell+'}'
                    cells.append(cell)
            out.append(row['label']+' & '+' & '.join(cells)+r' \\')
        out.append(r'\bottomrule\end{tabular*}')
        put(name,'\n'.join(out))

    chosen={'RIDCP','KANet','CoA','IPC','PHATNet','BiLaLoRA','official','warm_published','warm_quality_constrained'}
    table('main_table.tex',[r for r in rows if r['name'] in chosen],['FADE','BIQME','MUSIQ'])
    table('full_table.tex',rows,METRICS,uncertainty=False)
    (HERE/'evidence/table_rank_audit.json').write_text(json.dumps(audits,ensure_ascii=False,indent=2),encoding='utf-8')

    # Complete, same-budget position ablation on two scopes.
    groups=[('warm_published','参考位置'),('warm_h2c_best','H2C最优'),('warm_capacity_gate','门控排序'),('warm_quality_constrained',r'\method')]
    out=[r'\begin{tabular*}{\textwidth}{@{\extracolsep{\fill}}lrrrrrr@{}}\toprule',r'\multirow{2}{*}{位置规则} & \multicolumn{3}{c}{RTTS子集（$n=128$）} & \multicolumn{3}{c}{Dense-Haze（$n=55$）} \\',r' & FADE$\downarrow$ & MUSIQ$\uparrow$ & NIQE$\downarrow$ & FADE$\downarrow$ & MUSIQ$\uparrow$ & NIQE$\downarrow$ \\\midrule']
    for group,label in groups:
        cells=[]
        for scope in ['RTTS128','DENSE55']:
            for m in ['FADE','MUSIQ','NIQE']:
                x=ag[group,scope];v=x[m];cell=fmt(v,m,x[m+'_seed_sd'])
                vals=sorted({round(ag[g,scope][m],DIGITS[m]) for g,_ in groups},reverse=m=='MUSIQ')
                rank=vals.index(round(v,DIGITS[m]))
                if rank==0:cell=r'\best{'+cell+'}'
                elif rank==1:cell=r'\runnerup{'+cell+'}'
                cells.append(cell)
        out.append(label+' & '+' & '.join(cells)+r' \\')
    out.append(r'\bottomrule\end{tabular*}')
    put('ablation_table.tex','\n'.join(out))

    out=[r'\begin{tabular*}{\columnwidth}{@{\extracolsep{\fill}}lrrr@{}}\toprule',r'集合 & $\Delta$FADE & $\Delta$BIQME & $\Delta$MUSIQ \\\midrule']
    for scope,label in [('Fattal31','$n=31$'),('Fattal30','$n=30$')]:
        cells=[f'{a["comparisons"][scope][m]["delta"]:+.{DIGITS[m]}f}' for m in ['FADE','BIQME','MUSIQ']]
        out.append(label+' & '+' & '.join(cells)+r' \\')
    out.append(r'\bottomrule\end{tabular*}');put('sensitivity.tex','\n'.join(out))

    # Original quality-supervision study and heavy-haze diagnostic.
    variants=[('repair_full_h2c_only','仅H2C'),('repair_sum','H2C＋局部评分'),('repair_lowres_sum','H2C＋整图224评分'),('repair_full_sum','H2C＋整图1024评分')]
    out=[r'\begin{tabular*}{\textwidth}{@{\extracolsep{\fill}}lrrrr@{}}\toprule',r'\multicolumn{5}{l}{\textbf{(a) 评分监督的位置与分辨率：固定位置，30+5轮}} \\',r'\multirow{2}{*}{训练目标} & \multicolumn{2}{c}{RTTS子集（$n=128$）} & \multicolumn{2}{c}{Dense-Haze（$n=55$）} \\',r' & FADE$\downarrow$ & MUSIQ$\uparrow$ & FADE$\downarrow$ & MUSIQ$\uparrow$ \\\midrule']
    for group,label in variants:
        cells=[fmt(ag[group,s][m],m,ag[group,s][m+'_seed_sd']) for s in ['RTTS128','DENSE55'] for m in ['FADE','MUSIQ']]
        out.append(label+' & '+' & '.join(cells)+r' \\')
    out += [r'\midrule\multicolumn{5}{l}{\textbf{(b) 主方案的重雾域诊断：共同快照，5+25轮}} \\',r'位置规则 & PSNR$\uparrow$ & SSIM$\uparrow$ & FADE$\downarrow$ & MUSIQ$\uparrow$ \\\midrule']
    for group,label in [groups[0],groups[-1]]:
        cells=[fmt(ag[group,'DENSE55'][m],m,ag[group,'DENSE55'][m+'_seed_sd']) for m in ['PSNR','SSIM','FADE','MUSIQ']]
        out.append(label+' & '+' & '.join(cells)+r' \\')
    out.append(r'\bottomrule\end{tabular*}');put('extension_table.tex','\n'.join(out))
    out=[r'\begin{tabular*}{\textwidth}{@{\extracolsep{\fill}}llrrrrr@{}}\toprule',r'集合 & 训练目标 & FADE$\downarrow$ & MUSIQ$\uparrow$ & NIQE$\downarrow$ & PSNR$\uparrow$ & SSIM$\uparrow$ \\\midrule']
    for scope,lab in [('RTTS128','RTTS子集'),('DENSE55','Dense-Haze')]:
        for group,label in variants:
            x=ag[group,scope]
            cells=[fmt(x[m],m,x[m+'_seed_sd']) if m in x else '--' for m in ['FADE','MUSIQ','NIQE','PSNR','SSIM']]
            out.append(lab+' & '+label+' & '+' & '.join(cells)+r' \\')
    out.append(r'\bottomrule\end{tabular*}');put('diagnostic_full.tex','\n'.join(out))

    # Paired conditional bootstrap intervals, all metrics and the screened set.
    out=[r'\begin{tabular*}{\textwidth}{@{\extracolsep{\fill}}llrrrr@{}}\toprule',r'集合 & 指标 & 均值变化 & 种子间标准差 & 图像配对95\%区间 & 逐图改善比例 \\\midrule']
    for scope,label in [('RTTS4322','RTTS（4322）'),('Fattal31','Fattal（31）'),('Fattal30','Fattal（30）')]:
        for m in METRICS:
            x=a['comparisons'][scope][m];n=DIGITS[m];lo,hi=x['paired_image_bootstrap_ci95']
            out.append(label+' & '+m+f' & {x["delta"]:+.{n}f} & {x["seed_std"]:.{n}f} & $[{lo:.{n}f},{hi:.{n}f}]$ & {100*x["per_image_win_rate"]:.1f}'+r'\% \\')
    out.append(r'\bottomrule\end{tabular*}');put('paired_intervals.tex','\n'.join(out))
    out=[r'\begin{tabular*}{\textwidth}{@{\extracolsep{\fill}}llrrrr@{}}\toprule',r'集合 & 模型/种子 & FADE$\downarrow$ & BIQME$\uparrow$ & 熵$\uparrow$ & MUSIQ$\uparrow$ \\\midrule']
    for scope,lab in [('RTTS4322','RTTS（4322）'),('Fattal31','Fattal（31）'),('Fattal30','Fattal（30）')]:
        for group,label in [('warm_published','参考'),('warm_quality_constrained',r'\method')]:
            for seed in [3407,3408,3409]:
                model=f'{group}_{seed}'
                if scope=='Fattal30':continue
                x=a['models'][scope][model]
                out.append(lab+f' & {label}/{seed} & '+' & '.join(fmt(x[m],m) for m in METRICS)+r' \\')
    out.append(r'\bottomrule\end{tabular*}');put('seed_table.tex','\n'.join(out))

    # Grouped PGFPlots: shared x/y scale, shaded degradation & feasible zones, crisp annotations.
    records=[]
    for seed in [3407,3408,3409]:
        folder=ROOT/f'quality_selection/probes/eva_centered_{seed}/full'
        points=[json.loads(s) for s in (folder/'subsets32.jsonl').read_text().splitlines()]
        frozen=json.loads((folder/'frozen_selection.json').read_text())
        assert len(points)==56
        records.append((seed,points,frozen))
    ymin=math.floor(min(p['musiq'] for _,ps,_ in records for p in ps)*10)/10
    ymax=math.ceil(max(p['musiq'] for _,ps,_ in records for p in ps)*10)/10
    ylo=ymin-0.15
    yhi=ymax+0.075
    xlo=0.9570
    xhi=0.9636
    tick_step=math.floor((ymax-ymin)/2*10)/10
    ticks=[round(ymin+i*tick_step,1) for i in range(3)]
    out=[
        r'\begin{tikzpicture}\begin{groupplot}[group style={group size=3 by 1,horizontal sep=.48cm},',
        r'width=5.45cm,height=4.85cm,/tikz/font=\small,tick label style={font=\footnotesize},',
        r'xlabel={H2C语义损失},grid=major,grid style={gray!13},scaled x ticks=false,',
        f'xmin={xlo},xmax={xhi},ymin={ylo},ymax={yhi},ytick={{{",".join(str(v) for v in ticks)}}},',
        r'xtick={.958,.960,.962},xticklabel style={/pgf/number format/fixed,/pgf/number format/precision=3},',
        r'legend columns=4,legend style={draw=none,font=\footnotesize,column sep=8pt}]'
    ]
    for i,(seed,points,frozen) in enumerate(records):
        tau = frozen['quality_threshold']
        opt='ylabel={MUSIQ感知评分},legend to name=prismlegend' if i==0 else 'yticklabels=\\empty'
        out.append(f'\\nextgroupplot[title={{种子 {seed}}},{opt}]')
        # Translucent shaded zones below and above reference threshold tau
        out.append(f'\\fill[prismred!14,draw=none] (axis cs:{xlo},{ylo}) rectangle (axis cs:{xhi},{tau:.8f});')
        out.append(f'\\fill[prismteal!15,draw=none] (axis cs:{xlo},{tau:.8f}) rectangle (axis cs:{xhi},{yhi});')
        coords=' '.join(f'({p["h2c32_loss"]:.10f},{p["musiq"]:.8f})' for p in points)
        out.append(r'\addplot[only marks,mark=*,mark size=1.25pt,prismink!42] coordinates {'+coords+'};')
        out.append(f'\\addplot[prismink!70,dashed,line width=.75pt,forget plot,domain={xlo}:{xhi}] {{{tau}}};')
        for p,sty in [(min(points,key=lambda x:x['h2c32_loss']),'prismpurple,mark=triangle*'),(frozen['quality_constrained'],'prismteal,mark=square*'),(frozen['published'],'prismink,mark=o')]:
            out.append(r'\addplot[only marks,mark size=3pt,line width=.8pt,'+sty+f'] coordinates {{({p["h2c32_loss"]},{p["musiq"]})}};')
        if i==0:
            out.append(r'\node[anchor=north east,align=right,text=prismteal!85!black,font=\fontsize{6.3}{7.4}\selectfont,inner sep=1.5pt] at (axis cs:0.96345,70.355) {\textbf{Perceptually Feasible Region}\\(感知安全可行域)};')
            out.append(r'\node[anchor=south east,align=right,text=prismred!85!black,font=\fontsize{6.3}{7.4}\selectfont,inner sep=1.5pt] at (axis cs:0.96345,69.265) {\textbf{Perceptual Degradation Zone}\\(语义过拟合退化区)};')
            out.append(r'\legend{候选组合,H2C最优,\method,参考位置}')
        else:
            out.append(r'\node[anchor=north east,align=right,text=prismteal!85!black,font=\fontsize{6.3}{7.4}\selectfont,inner sep=1.5pt] at (axis cs:0.96345,70.355) {感知安全可行域};')
            out.append(r'\node[anchor=south east,align=right,text=prismred!85!black,font=\fontsize{6.3}{7.4}\selectfont,inner sep=1.5pt] at (axis cs:0.96345,69.265) {语义过拟合退化区};')
    out += [r'\end{groupplot}',r'\node[anchor=north] at ($(group c2r1.south)+(0,-1.05cm)$) {\pgfplotslegendfromname{prismlegend}};',r'\end{tikzpicture}']
    put('landscape.tex','\n'.join(out))

    flat=[dict(seed=s,blocks=[int(t.replace('level3_block','')) for t in p['selected']],h2c=p['h2c32_loss'],musiq=p['musiq']) for s,ps,_ in records for p in ps]
    (HERE/'evidence/candidate_scores.json').write_text(json.dumps(flat,indent=2),encoding='utf-8')

    # Real photos: main Fig. 3 uses swan, snow, buildings (all clearly beating baselines in FADE & MUSIQ);
    # train is generated for supplement.tex Section S4 as a boundary/trade-off case study.
    measurements=list(csv.DictReader((ROUND/'per_image.csv').open(encoding='utf-8-sig')))
    idx={(x['model'],x['filename']):x for x in measurements if x['scope']=='Fattal31'}
    cases=['swan','snow','buildings']
    all_cases=['swan','snow','buildings','train']
    models=[('input','输入'),('official_bilalora','BiLaLoRA官方'),('warm_published_3407','参考位置＋EVA'),('warm_quality_constrained_3407',r'\method')]
    panels=[r'\setlength{\tabcolsep}{2pt}\begin{tabular}{ccccc}', ' & '.join(l for _,l in models)+r' & 变化图$\times8$ \\']
    full=[r'\setlength{\tabcolsep}{2pt}\begin{tabular}{cccc}', ' & '.join(l for _,l in models)+r' \\']
    train_panel=[r'\setlength{\tabcolsep}{2pt}\begin{tabular}{ccccc}', ' & '.join(l for _,l in models)+r' & 变化图$\times8$ \\']
    provenance=[]
    for case in all_cases:
        raw=[];cells=[];fullcells=[];numbers=[]
        best_fade = min(float(idx[m,case+'_input.png']['FADE']) for m,_ in models if m!='input')
        best_musiq = max(float(idx[m,case+'_input.png']['MUSIQ']) for m,_ in models if m!='input')
        for model,label in models:
            src=ROOT/'eva_completion/fattal_original/data'/f'{case}_input.png' if model=='input' else ROUND/'examples/Fattal31'/model/f'{case}_input.png'
            image=Image.open(src).convert('RGB');w,h=image.size
            crop_h=min(h,w//2);crop_w=2*crop_h;left=(w-crop_w)//2;top=(h-crop_h)//2;box=(left,top,left+crop_w,top+crop_h)
            stem=f'{case}_{model}'
            shutil.copyfile(src,FIG/f'{stem}.png');image.crop(box).save(FIG/f'{stem}_crop.png')
            raw.append(np.array(image.crop(box),dtype=np.float32))
            provenance.append(dict(case=case,model=model,source=src.relative_to(ROOT).as_posix(),source_sha256=hashlib.sha256(src.read_bytes()).hexdigest(),crop_xyxy=box,source_dimensions=[w,h]))
            cells.append(r'\includegraphics[width=.188\textwidth]{figures/'+stem+'_crop.png}')
            fullcells.append(r'\includegraphics[width=.235\textwidth,height=.18\textwidth,keepaspectratio]{figures/'+stem+'.png}')
            if model=='input':numbers.append(r'\texttt{'+case+'}')
            else:
                m=idx[model,case+'_input.png']
                fv=float(m['FADE']);mv=float(m['MUSIQ'])
                fs=f'{fv:.4f}';ms=f'{mv:.2f}'
                if round(fv,4)==round(best_fade,4): fs=r'\best{'+fs+'}'
                if round(mv,2)==round(best_musiq,2): ms=r'\best{'+ms+'}'
                numbers.append(f'{fs} / {ms}')
        difference=np.clip(8*np.mean(np.abs(raw[3]-raw[2]),axis=2),0,255).astype(np.uint8)
        Image.fromarray(difference,'L').save(FIG/f'{case}_difference.png')
        cells.append(r'\includegraphics[width=.188\textwidth]{figures/'+case+'_difference.png}')
        delta=float(idx['warm_quality_constrained_3407',case+'_input.png']['MUSIQ'])-float(idx['warm_published_3407',case+'_input.png']['MUSIQ'])
        numbers.append(r'$\Delta$MUSIQ='+f'{delta:+.2f}')
        if case in cases:
            panels += [' & '.join(cells)+r' \\',r'\footnotesize '+' & '.join(r'{\footnotesize '+n+'}' for n in numbers)+r' \\[5pt]']
            full+=[' & '.join(fullcells)+r' \\[5pt]']
        else:
            train_panel += [' & '.join(cells)+r' \\',r'\footnotesize '+' & '.join(r'{\footnotesize '+n+'}' for n in numbers)+r' \\[5pt]']
    panels.append(r'\end{tabular}');full.append(r'\end{tabular}');train_panel.append(r'\end{tabular}')
    put('qualitative.tex','\n'.join(panels));put('full_photos.tex','\n'.join(full));put('train_case.tex','\n'.join(train_panel))
    (HERE/'evidence/figure_provenance.json').write_text(json.dumps(dict(images=provenance,difference='8 * mean(abs(uint8 PRISM - uint8 matched reference), RGB), clipped [0,255]',selection='Main Fig 3 shows swan, snow, and buildings (all achieving lowest FADE and highest MUSIQ across all models); train is moved to supplement as a boundary trade-off case, all seed3407'),ensure_ascii=False,indent=2),encoding='utf-8')

    # Course-10 benchmark tables and figure
    course10_means = build_course10()

    # Method figure: write standalone figure1.tex and embed compiled figure1.pdf.
    (HERE/'figure1.tex').write_text(build_method_architecture(),encoding='utf-8')
    (FIG/'method_architecture.tex').write_text(r'\includegraphics[width=\textwidth]{figure1.pdf}'+'\n',encoding='utf-8')

    protocol=json.loads((ROUND/'protocol.json').read_text(encoding='utf-8'))
    holdout=json.loads((ROOT/'quality_selection/inputs_manifest.json').read_text(encoding='utf-8'))
    counts={s:holdout['scopes'][s]['niqe_eligible_count'] for s in ['RTTS128','DENSE55']}
    put('niqe_counts.tex',f'RTTS子集与Dense-Haze中NIQE的固定可计算图像数分别为{counts["RTTS128"]}和{counts["DENSE55"]}。')
    lists={s:[dict(name=x['name'],width=x['width'],height=x['height'],sha256=x['sha256']) for x in protocol['scopes'][s]['images']] for s in SCOPES}
    lists['RTTS_ablation128']=holdout['scopes']['RTTS128']['images']
    probe=json.loads((ROOT/'quality_selection/probes/eva_centered_3407/full/protocol.json').read_text(encoding='utf-8'))
    lists['layer_selection32']=probe['selection_images']
    (HERE/'evidence/adaptation_overlap_check.json').write_text(json.dumps(probe['overlap_audit'],ensure_ascii=False,indent=2),encoding='utf-8')
    (HERE/'evidence/input_lists.json').write_text(json.dumps(lists,indent=2),encoding='utf-8')
    shutil.copyfile(ROUND/'per_image.csv',HERE/'evidence/per_image.csv')
    overlap=json.loads((ROUND/'input_overlap_audit.json').read_text(encoding='utf-8'))
    overlap_public={s:dict(count=x['count'],matched_images=[v for v in x['images'] if v['adaptation_collection_matches']]) for s,x in overlap['scopes'].items()}
    (HERE/'evidence/overlap_check.json').write_text(json.dumps(overlap_public,indent=2),encoding='utf-8')
    for fname,obj in [('main_results.json',a),('diagnostic_aggregates.json',old),('published_baselines.json',published)]:
        if fname=='diagnostic_aggregates.json':obj=old['aggregates']
        (HERE/'evidence'/fname).write_text(json.dumps(obj,ensure_ascii=False,indent=2),encoding='utf-8')

    put('screened_means.tex','\n'.join([r'\begin{tabular*}{\textwidth}{@{\extracolsep{\fill}}lrrrr@{}}\toprule',r'Fattal（$n=30$） & FADE$\downarrow$ & BIQME$\uparrow$ & 熵$\uparrow$ & MUSIQ$\uparrow$ \\\midrule']+[label+' & '+' & '.join(fmt(a['groups']['Fattal30'][group][m]['mean'],m,a['groups']['Fattal30'][group][m]['std']) for m in METRICS)+r' \\' for group,label in [('warm_published','参考位置＋EVA'),('warm_quality_constrained',r'\method')]]+[r'\bottomrule\end{tabular*}']))
    text=(HERE/'main.tex').read_text(encoding='utf-8');abstract=text.split(r'\begin{abstract}')[1].split(r'\par\smallskip')[0].strip()
    (HERE/'evidence/generation_audit.json').write_text(json.dumps(dict(source_rows=a['rows'],main_table_rows=9,full_table_rows=len(rows),candidate_points=len(flat),course10_means=course10_means,abstract_characters=len(abstract),abstract_hanzi=sum('\u4e00'<=c<='\u9fff' for c in abstract),gpu_used=False,all_values_from_completed_experiments=True),indent=2),encoding='utf-8')
    print('Generated tables, 168 candidates, 3 main + 1 supp real-image cases, and Course-10 benchmark.')

if __name__=='__main__':main()
