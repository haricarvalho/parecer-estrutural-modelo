#!/usr/bin/env python3
"""Monta o Parecer Estrutural Motiva a partir do modelo (modelo_parecer.docx) e de um JSON.
Uso: python montar_parecer.py modelo_parecer.docx entrada.json saida.docx
JSON: {objeto, solicitante, demanda, unidade, localizacao, fase, setor, emissor, data, codigo, revisao,
 status (Aprovado|Aprovado com ressalvas|Reprovado), justificativa, documentos,
 comentarios:[{titulo, texto, figura_legenda?, figura_png?}]}
Regras: sem comentarios => status Aprovado; legenda de figura ate 220 caracteres."""
import sys, json, zipfile, copy, re, shutil, os
from lxml import etree
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ns={'w':W}
q=lambda t:'{%s}%s'%(W,t)
CORES={'Aprovado':('00B050','E2EFDA'),'Aprovado com ressalvas':('BF9000','FFF2CC'),'Reprovado':('C00000','F8CBAD')}

def set_text(sdt,valor):
    c=sdt.find('w:sdtContent',ns)
    runs=c.findall('.//w:r',ns)
    ts=c.findall('.//w:t',ns)
    ts[0].text=valor; ts[0].set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    for t in ts[1:]: t.text=''
def main(modelo,entrada,saida):
    d=json.load(open(entrada,encoding='utf8'))
    com=d.get('comentarios',[])
    if not com: d['status']='Aprovado'
    assert d['status'] in CORES, 'status invalido'
    zin=zipfile.ZipFile(modelo)
    doc=etree.fromstring(zin.read('word/document.xml'))
    body=doc.find('w:body',ns)
    campos={'Objeto da análise':'objeto','Solicitante':'solicitante','Demanda':'demanda','Unidade de Negócio':'unidade','Localização':'localizacao','fase_projeto':'fase','Setor de análise':'setor','Emissor':'emissor','Data':'data','codigo_parecer':'codigo','Justificativa do status':'justificativa','Documentos analisados':'documentos'}
    for s in body.iter(q('sdt')):
        al=s.find('w:sdtPr/w:tag',ns).get(q('val'))
        if al in campos and campos[al] in d: set_text(s,str(d[campos[al]]))
        elif al=='revisao_parecer' and 'revisao' in d: set_text(s,d['revisao'])
        elif al=='Status da análise':
            set_text(s,d['status']); cor,fundo=CORES[d['status']]
            rpr=s.find('.//w:r/w:rPr',ns)
            rpr.find('w:color',ns).set(q('val'),cor); rpr.find('w:shd',ns).set(q('fill'),fundo)
    # blocos prototipo dos comentarios
    ch=list(body)
    idx=[i for i,e in enumerate(ch) if ''.join(e.itertext()).strip()=='Comentários'][0]
    proto_c=ch[idx+1]; proto_i=ch[idx+2]; proto_l=ch[idx+3]
    sect=ch[-1]
    for e in ch[idx+1:-1]: body.remove(e)
    media={}; rels_extra=[]
    zin_rels=etree.fromstring(zin.read('word/_rels/document.xml.rels'))
    novos=[]
    for n,c in enumerate(com,1):
        p=copy.deepcopy(proto_c); rs=p.findall('w:r',ns)
        rs[0].find('w:t',ns).text=c['titulo'].rstrip()+('' if c['titulo'].rstrip().endswith('.') else '.')+' '
        rs[1].find('w:t',ns).text=c['texto']
        for r in rs[2:]: p.remove(r)
        body.insert(list(body).index(sect),p)
        leg=c.get('figura_legenda')
        if leg:
            assert len(leg)<=220,'legenda >220'
            if c.get('figura_png'):
                pi=copy.deepcopy(proto_i); rid='rIdFig%d'%n
                pi.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip').set('{%s}embed'%R,rid)
                from PIL import Image
                w,h=Image.open(c['figura_png']).size; cx=5715000; cy=int(cx*h/w)
                for e in pi.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent','{http://schemas.openxmlformats.org/drawingml/2006/main}ext'):
                    e.set('cx',str(cx)); e.set('cy',str(cy))
                pi.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr').set('id',str(100+n))
                body.insert(list(body).index(sect),pi); media[rid]=c['figura_png']
            pl=copy.deepcopy(proto_l); ts=pl.findall('.//w:t',ns)
            ts[0].text='Figura '; ts[1].text=str(n); ts[2].text=': '+leg.split(':',1)[-1].strip() if leg.startswith('Figura') else ': '+leg
            body.insert(list(body).index(sect),pl)
    # rels e midia
    rels=zin_rels
    RN='http://schemas.openxmlformats.org/package/2006/relationships'
    for rid,path in media.items():
        e=etree.SubElement(rels,'{%s}Relationship'%RN); e.set('Id',rid); e.set('Type',R+'/image'); e.set('Target','media/'+rid+'.png')
    zout=zipfile.ZipFile(saida,'w',zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        if it.filename=='word/document.xml': data=etree.tostring(doc,xml_declaration=True,encoding='UTF-8',standalone=True)
        elif it.filename=='word/_rels/document.xml.rels': data=etree.tostring(rels,xml_declaration=True,encoding='UTF-8',standalone=True)
        else: data=zin.read(it.filename)
        zout.writestr(it,data)
    for rid,path in media.items(): zout.write(path,'word/media/'+rid+'.png')
    zout.close(); print('ok',saida)
if __name__=='__main__': main(*sys.argv[1:4])
