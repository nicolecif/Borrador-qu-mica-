"""Gr谩ficas del proyecto QUIMORG a partir de compuestos(1).xlsx. Uso: python graficas.py "compuestos(1).xlsx" Dependencias: pandas, openpyxl, numpy, matplotlib. No modifica el Excel. Exporta una tabla unificada CSV y figuras PNG/PDF. """
import argparse
from pathlib import Path
import re
import unicodedata
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

PROPS = ['Peso molecular (g/mol)', 'Fusi贸n (掳C)', 'Ebullici贸n (掳C)',
         'Densidad (g/cm鲁)', 'pKa', 'Momento dipolar (D)',
         'Solubilidad (g/L)', 'Presi贸n de vapor (Pa)']
ALCANOS = {'metano', 'etano', 'propano', 'butano', 'isobutano'}
ALQUENOS = {'eteno', 'propeno', 'but-1-eno', 'but-2-eno', 'pent-1-eno'}
ROLES = ['Reacci贸n', 'Obtenci贸n']

# Clasificaci贸n descriptiva por palabras iniciales o expresiones expl铆citas.
# Se conserva el texto 铆ntegro de la fuente en la tabla unificada.
REGLAS = [
    ('Despolimerizaci贸n', r'despolimer'),
    ('Polimerizaci贸n', r'polimerizacion|copolimer'),
    ('Deshidrogenaci贸n', r'deshidrogenacion'),
    ('Hidrogenaci贸n', r'(?:semi)?hidrogenacion'),
    ('Hidrogen贸lisis / craqueo', r'hidrogenolisis|craqueo|reformado'),
    ('Oxidaci贸n', r'oxidacion|epoxidacion|amoxidacion|dihidroxilacion|acetoxilacion oxidativa'),
    ('Reducci贸n', r'reduccion|hidrodesoxigenacion|desulfuracion reductora|hidrodescloracion'),
    ('Halogenaci贸n', r'halogenacion|oxicloracion'),
    ('Eliminaci贸n', r'eliminacion|deshidrohalogenacion|deshalogenacion|deshidratacion'),
    ('Alquilaci贸n', r'alquilacion'),
    ('Acoplamiento', r'acoplamiento'),
    ('Carbonilaci贸n / hidroformilaci贸n', r'carbonilacion|hidroformilacion'),
    ('Adici贸n', r'adicion|hidratacion|hidroboracion|hidrozirconacion|hidrotiolacion'),
    ('Sustituci贸n', r'sustitucion'),
    ('Isomerizaci贸n', r'isomerizacion'),
    ('Dimerizaci贸n', r'dimerizacion'),
    ('Met谩tesis', r'metatesis'),
    ('Carbenos', r'insercion de carbeno|ciclopropanacion'),
    ('Coordinaci贸n', r'formacion de complejo|coordinacion'),
]

def texto(valor):
    if pd.isna(valor):
        return ''
    return ''.join(c for c in unicodedata.normalize('NFD', str(valor).strip().lower())
                   if unicodedata.category(c) != 'Mn')

def tipo_general(valor):
    t = texto(valor)
    if not t or t.startswith('no reportado'):
        return 'No reportado'
    for nombre, patron in REGLAS:
        if re.match(patron, t):
            return nombre
    return 'Otros tipos'

def medio_general(valor):
    t = texto(valor)
    if not t or re.search(r'no determinad|no indicado|no especificad|no reportado', t):
        return 'Otro / no determinado'
    for palabra, nombre in [('acido','脕cido'), ('basico','B谩sico'), ('neutro','Neutro'), ('apolar','Apolar')]:
        if re.match(palabra, t):
            return nombre
    # Solo se agrupan descripciones inequ铆vocas sin inferir el medio del solvente.
    if 'medio acido' in t or 'medio catalitico acido' in t:
        return '脕cido'
    return 'Otro / no determinado'

NUM = r'[+-]?\d+(?:[.,]\d+)?(?:\s*[x脳]\s*10\s*\^\s*[+-]?\d+|[eE][+-]?\d+)?'
SUP = str.maketrans('鈦奥孤猜斥伌鈦碘伓鈦封伕鈦光伝鈦�','0123456789-+')

def normalizar(v):
    s = str(v).strip().replace('鈭�','-').replace('\xa0',' ')
    return re.sub(r'[鈦奥孤猜斥伌鈦碘伓鈦封伕鈦光伝鈦篯+',lambda m:'^'+m.group().translate(SUP),s)

def numero(v):
    return float(re.sub(r'\s*[x脳]\s*10\s*\^\s*','e',v.replace(',','.')).replace(' ',''))

def temperatura(v):
    if pd.isna(v): return np.nan
    if isinstance(v,(int,float,np.number)): return float(v)
    # Solo cifra o intervalo sencillo, no etapas, cotas, ambiente o notas mixtas.
    s=normalizar(v)
    m=re.fullmatch(rf'({NUM})(?:\s*[鈥撯€�-]\s*({NUM}))?\s*(?:掳\s*C)?',s)
    if not m: return np.nan
    a=numero(m[1]); b=numero(m[2]) if m[2] else a
    return (a+b)/2

def propiedad(v, p):
    if pd.isna(v): return np.nan
    # Preservar casos cualitativos, desigualdades y rangos en el texto fuente.
    s=normalizar(v)
    m=re.match(rf'^({NUM})\s*(.*)$',s)
    if not m: return np.nan
    n=numero(m[1]); resto=m[2].strip()
    if re.match(r'^[鈥撯€�-]|^;|^a\s',resto): return np.nan
    if p=='Presi贸n de vapor (Pa)':
        if ';' in resto or 'tambi茅n' in resto.lower(): return np.nan
        u=re.match(r'^(MPa|kPa|hPa|Pa|mmHg|Torr|atm|bar|psi)\b',resto)
        if not u: return np.nan
        n*= {'MPa':1e6,'kPa':1e3,'hPa':100,'Pa':1,'mmHg':101325/760,'Torr':101325/760,'atm':101325,'bar':1e5,'psi':6894.757}[u[1]]
    elif p=='Solubilidad (g/L)':
        u=re.match(r'^(mg/mL|mg/L|g/L)\b',resto)
        if not u: return np.nan
        n*= {'mg/mL':1,'mg/L':.001,'g/L':1}[u[1]]
        # ppm y otras unidades no se estiman.
    elif p=='Densidad (g/cm鲁)':
        if resto and not re.match(r'^g/(?:cm\^?3|mL)\b',resto): return np.nan
    elif p=='Ebullici贸n (掳C)' or p=='Fusi贸n (掳C)':
        if resto and not re.match(r'^(?:掳\s*C|\()',resto): return np.nan
    elif resto and not re.match(r'^(?:g/mol|D\b|\()',resto):
        return np.nan
    return n

def cargar(archivo):
    hojas=pd.read_excel(archivo,sheet_name=None,header=None)
    registros=[]
    for compuesto, hoja in hojas.items():
        familia='Alcanos' if compuesto.lower() in ALCANOS else 'Alquenos' if compuesto.lower() in ALQUENOS else 'Otro'
        # Localizaci贸n de los encabezados por sus nombres, no por posiciones fijas.
        i=hoja.index[hoja[0].astype(str).str.strip().eq('Clase de registro')][0]
        j=hoja.index[hoja[0].astype(str).str.strip().eq('Peso molecular (g/mol)')][0]
        base=dict(zip(hoja.iloc[j,:8],hoja.iloc[j+1,:8]))
        fila={'Compuesto':compuesto,'Familia':familia,'Clase de registro':'Compuesto','Sustancia relacionada':compuesto,'Hoja':compuesto,'Fila Excel':int(j+2)}
        fila.update({p+' original':base.get(p) for p in PROPS})
        registros.append(fila)
        tabla=hoja.iloc[i+1:].copy(); tabla.columns=hoja.iloc[i]
        tabla=tabla.loc[:,tabla.columns.notna()]
        propiedades=tabla.columns[-8:]
        for idx,r in tabla.iterrows():
            rol=r['Clase de registro']
            if rol not in ROLES: continue
            fila=r.to_dict(); fila.update({'Compuesto':compuesto,'Familia':familia,'Hoja':compuesto,'Fila Excel':int(idx+1),
                    'Sustancia relacionada': r['Producto principal'] if rol=='Reacci贸n' else r['Reactante evaluado']})
            for p,col in zip(PROPS,propiedades): fila[p+' original']=r[col]
            registros.append(fila)
    datos=pd.DataFrame(registros)
    datos.insert(0,'ID fila',range(1,len(datos)+1))
    for p in PROPS: datos[p]=datos[p+' original'].map(lambda v:propiedad(v,p))
    datos['Tipo general']=datos['Tipo de reacci贸n'].map(tipo_general)
    datos['Medio general']=datos['Medio'].map(medio_general)
    datos['Temperatura central (掳C)']=datos['Temperatura (掳C)'].map(temperatura)
    return datos

def graficar(datos,salida):
    salida.mkdir(parents=True,exist_ok=True)
    figuras=salida/'graficas'; figuras.mkdir(exist_ok=True)
    tablas=salida/'tablas'; tablas.mkdir(exist_ok=True)
    plt.rcParams.update({'font.size':11,'axes.spines.top':False,'axes.spines.right':False,'figure.dpi':120})
    rx=datos[datos['Clase de registro'].isin(ROLES)].copy()
    comp=datos[datos['Clase de registro']=='Compuesto'].copy()
    def guardar(fig,nombre):
        fig.tight_layout(); fig.savefig(figuras/(nombre+'.png'),dpi=260,bbox_inches='tight'); fig.savefig(figuras/(nombre+'.pdf'),bbox_inches='tight'); plt.close(fig)
    def exportar(t,n): t.to_csv(tablas/(n+'.csv'),encoding='utf-8-sig')
    def barras(t,titulo,nombre):
        t=t.loc[t.sum(axis=1).sort_values().index]; fig,ax=plt.subplots(figsize=(10,max(4,0.40*len(t)+2)))
        t.plot.barh(ax=ax,stacked=True); ax.set(title=titulo,xlabel='N煤mero de registros',ylabel=''); ax.legend(title=''); guardar(fig,nombre); exportar(t,nombre)
    # 1. Gr谩fico circular: familias en registros de reactividad.
    f=rx['Familia'].value_counts(); fig,ax=plt.subplots(figsize=(7,5)); ax.pie(f,labels=f.index,autopct='%1.1f%%',startangle=90); ax.set_title(f'Registros de reactividad por familia (n={len(rx)})'); guardar(fig,'01_familias'); exportar(f,'01_familias')
    # 2. Anillo de clases de registro.
    f=rx['Clase de registro'].value_counts(); fig,ax=plt.subplots(figsize=(7,5)); ax.pie(f,labels=f.index,autopct='%1.1f%%',wedgeprops={'width':.4},startangle=90); ax.set_title(f'Reacciones y obtenciones (n={len(rx)})'); guardar(fig,'02_reacciones_obtenciones'); exportar(f,'02_reacciones_obtenciones')
    # 3. Frecuencia de categor铆as de reacci贸n, por rol.
    barras(pd.crosstab(rx['Tipo general'],rx['Clase de registro']), 'Tipos de reacci贸n por clase de registro','03_tipos_reaccion')
    # 4. Frecuencias por compuesto.
    t=pd.crosstab(rx['Compuesto'],rx['Tipo general']).reindex(comp['Compuesto']); fig,ax=plt.subplots(figsize=(15,6)); im=ax.imshow(t,aspect='auto',cmap='Blues'); ax.set_xticks(range(len(t.columns)),t.columns,rotation=55,ha='right'); ax.set_yticks(range(len(t)),t.index); ax.set_title('Tipos de reacci贸n por compuesto');
    for i in range(len(t)):
        for j in range(len(t.columns)):
            if t.iat[i,j]: ax.text(j,i,str(t.iat[i,j]),ha='center',va='center',fontsize=9,color='white' if t.iat[i,j]>t.to_numpy().max()/2 else 'black')
    fig.colorbar(im,ax=ax,label='Registros'); guardar(fig,'04_mapa_calor'); exportar(t,'04_mapa_calor')
    # 5. Medio de reacci贸n.
    barras(pd.crosstab(rx['Medio general'],rx['Clase de registro']),'Medios de reacci贸n (clasificaci贸n por texto)','05_medios')
    # 6. Tipos de reactividad separados por familia.
    barras(pd.crosstab(rx['Tipo general'],rx['Familia']),'Tipos de reacci贸n por familia qu铆mica','06_familias_tipos')
    # 7. Temperaturas: puntos por familia y cajas con datos v谩lidos.
    temp='Temperatura central (掳C)'; t=rx.dropna(subset=[temp]); fig,ax=plt.subplots(figsize=(9,5)); grupos=list(t.groupby('Familia')); ax.boxplot([g[temp].to_numpy() for _,g in grupos],showfliers=False);
    for i,(nombre,g) in enumerate(grupos,1): ax.scatter(np.full(len(g),i)+np.linspace(-.10,.10,len(g)),g[temp],alpha=.45,s=18)
    ax.set_xticks(range(1,len(grupos)+1),[f'{n}\n(n={len(g)})' for n,g in grupos]); ax.set(ylabel='Temperatura (掳C)',title='Temperatura en valores e intervalos simples'); guardar(fig,'07_temperatura'); exportar(t.groupby('Familia')[temp].agg(['count','min','median','max']),'07_temperatura')
    # 8. Masa vs ebullici贸n: solo compuestos base, sin ajustar regresi贸n.
    x,y=PROPS[:3:2]; t=comp.dropna(subset=[x,y]); fig,ax=plt.subplots(figsize=(9,6));
    for nombre,g in t.groupby('Familia'):
        ax.scatter(g[x],g[y],label=nombre,s=65)
        for k,(_,r) in enumerate(g.iterrows()): ax.annotate(r['Compuesto'],(r[x],r[y]),xytext=(6,(-12 if k%2 else 7)),textcoords='offset points',fontsize=9)
    ax.set(xlabel=x,ylabel=y,title='Masa molecular y punto de ebullici贸n reportado'); ax.legend(); ax.margins(.15); guardar(fig,'08_masa_ebullicion'); exportar(t[['Compuesto','Familia',x,y]],'08_masa_ebullicion')
    # 9. Disponibilidad num茅rica de las propiedades (no equivalencia experimental).
    t=rx.groupby('Clase de registro')[PROPS].count().T; barras(t,'Disponibilidad de valores num茅ricos interpretables','09_disponibilidad')
    # 10. Burbujas: frecuencia por categor铆a y familia, no medidas qu铆micas.
    t=rx.groupby(['Tipo general','Familia']).size().rename('Registros').reset_index(); cats=t.groupby('Tipo general')['Registros'].sum().sort_values().index.tolist(); fig,ax=plt.subplots(figsize=(9,max(5,len(cats)*.4))); 
    for i,(fam,g) in enumerate(t.groupby('Familia')): ax.scatter([i]*len(g),g['Tipo general'].map(cats.index),s=g['Registros']*24,alpha=.65,label=fam)
    ax.set_xticks(range(t['Familia'].nunique()),sorted(t['Familia'].unique())); ax.set_yticks(range(len(cats)),cats); ax.set(title='Frecuencia por tipo y familia (谩rea proporcional al conteo)',ylabel=''); ax.set_xlim(-.6,t['Familia'].nunique()-.4); guardar(fig,'10_burbujas'); exportar(t,'10_burbujas')
    datos.to_csv(salida/'tabla_unificada.csv',index=False,encoding='utf-8-sig')
    print('Compuestos:',len(comp),'Reacciones:',sum(rx['Clase de registro']=='Reacci贸n'),'Obtenciones:',sum(rx['Clase de registro']=='Obtenci贸n'))
    print('Temperaturas simples:',rx[temp].notna().sum(),'de',len(rx))
    print('Resultados:',salida.resolve())
    (salida/'LEEME.txt').write_text('Las figuras se generan a partir del Excel indicado. Las categor铆as generales son aproximaciones por texto; se conserva el original. Las temperaturas solo se resumen si son valores o intervalos simples; no se imputan datos. Las propiedades con condiciones diferentes no se consideran experimentalmente equivalentes. Las figuras son est谩ticas; los gr谩ficos nativos del PBIP se construyen en Power BI.\n',encoding='utf-8')

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('excel',nargs='?',default='compuestos(1).xlsx')
    parser.add_argument('--salida',default='resultados_graficas')
    args=parser.parse_args(); datos=cargar(Path(args.excel)); graficar(datos,Path(args.salida))

if __name__=='__main__': main()
