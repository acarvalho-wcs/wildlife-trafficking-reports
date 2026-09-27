from pathlib import Path
import json
import re

ROOT = Path(".")
BASE_SCRIPT = ROOT / ".github/scripts/backfill_week_2026_09_08_14.py"
BASE_LANGS = ROOT / ".github/data/backfill_LANGS.json"
TEMP_LANGS = ROOT / ".github/data/_week_2026_09_14_20_langs.json"
PERIOD_ID = "2026-09-14_2026-09-20"

langs = json.loads(BASE_LANGS.read_text(encoding="utf-8"))

updates = {
"pt": {
"outfile":"Relatorio_Semanal_Observatorio_Global_14-20_Set_2026_WCS.docx",
"title":"RELATÓRIO SEMANAL DE MONITORAMENTO",
"subtitle":"Observatório Global de Tráfico de Animais",
"period":"14-20 de setembro de 2026 | Semana civil | America/Manaus",
"close":"WCS Brasil | Programa de Inteligência para Conservação",
"metrics":["25\ncasos validados","13\npaíses/territórios","680\nanimais/espécimes em contagens exatas",">=21\nadicionais em contagens mínimas"],
"foot":"Publicado em 27 de setembro de 2026. O núcleo semanal usa a data de ocorrência/apreensão, e não apenas a data de publicação.",
"scope":"Núcleo semanal: somente registros VALIDATED com ocorrência/apreensão entre segunda-feira 14/09/2026 00:00 e domingo 20/09/2026 23:59:59, America/Manaus. Quantidades mínimas e desconhecidas não entram no total exato.",
"sections":["1. Resumo executivo","2. Núcleo semanal consolidado","3. Leitura analítica da semana","4. Casos do período","5. Limites da evidência","6. Nota metodológica","7. Fontes principais"],
"executive":[
"O núcleo semanal reúne 25 casos VALIDATED em 13 países ou territórios. Dezenove casos têm contagens exatas, somando 680 animais ou espécimes. Dois casos adicionais apresentam somente contagens mínimas, com ao menos 21 indivíduos confirmados; quatro casos permanecem sem quantidade numérica confiável.",
"Fatos documentados: o maior incidente por contagem exata envolve 206 aves na BR-222, em São Gonçalo do Amarante. Dois incidentes registraram 101 espécimes cada, em Singapura e Medan. Em Samalayuca foram apreendidas 75 tarântulas e 20 tartarugas; em Motu, 55 tartarugas-de-casco-mole-indianas.",
"Foram documentados canais logísticos distintos - rodovias, ônibus, encomenda postal, bagagem em conexão aeroportuária, encomenda aérea e porto marítimo - além de exploração ou caça ilegal sem rota comercial demonstrada."
],
"analytical":[
["Padrões observados","O transporte terrestre aparece como modal principal em 8 dos 25 registros; 2 são classificados como aéreos e 1 como marítimo. Em 14 registros, o modal principal não foi informado. Aves e répteis aparecem em grande parte do núcleo; mamíferos também estão presentes em casos de grandes felinos, tatus, pangolins e ouriços. Há recorrência de transição do ambiente digital para a apreensão física, como WhatsApp em Angaco, anúncios online em Singapura e investigação de vendas por redes sociais em Mairiporã."],
["Interpretação analítica","Os casos mostram um mosaico de cadeias de suprimento e exploração que utilizam logística comercial, bagagem de passageiros, transporte rodoviário e canais digitais. Isso reforça o valor de combinar controles em nós de transporte, inteligência digital e rastreabilidade documental. A recorrência de táxons distintos em canais semelhantes sugere que indicadores de risco devem priorizar comportamento, rota, documentação e método de ocultação, e não apenas espécies-alvo. Essas semelhanças não demonstram, por si só, uma rede transnacional comum."],
["Limites da evidência","Quatro casos não têm quantidade confiável e dois apresentam somente contagens mínimas; portanto, 680 não representa todos os animais envolvidos. As fontes de Motu divergem em um dia, mas ambas as datas estão na janela semanal. Em Singapura, a autoridade não afirmou que as seis pessoas investigadas formassem uma única rede. Em Paulo Afonso, a direção Bahia-Pernambuco não é estabelecida com segurança. Origem biológica, comprador final, rota completa e participação em redes permanecem desconhecidos em vários registros."]
],
"methodology":[
"Inclusão definida pela data de ocorrência/apreensão entre 14 e 20 de setembro de 2026, no fuso America/Manaus; a data de publicação foi usada apenas como informação auxiliar.",
"Somente registros VALIDATED do banco canônico do Observatório integram o núcleo semanal. Não foram incluídos casos cuja data do evento não pôde ser estabelecida dentro da janela.",
"Operações agregadas e comunicados conjuntos foram reconciliados para evitar dupla contagem. Hoedspruit e Gravelotte são duas intervenções distintas, cada uma com um pangolim; prisões e veículos divulgados conjuntamente não foram repartidos entre os dois registros.",
"Peso, valor financeiro, embalagens, partes e produtos não foram convertidos automaticamente em número de animais. Volta Redonda permanece sem quantidade por inconsistência interna da fonte."
],
"core_intro":"Ordem cronológica pela data da ocorrência/apreensão. Quantidades mínimas ou desconhecidas são excluídas do total exato de 680.",
"body_footer":"Criado e gerenciado pelo Programa de Inteligência para Conservação da WCS Brasil."
},
"en": {
"outfile":"Weekly_Report_Global_Observatory_14-20_Sep_2026_WCS.docx",
"title":"WEEKLY MONITORING REPORT",
"subtitle":"Illegal Wildlife Trafficking Global Observatory",
"period":"14-20 September 2026 | Civil week | America/Manaus",
"close":"WCS Brasil | Conservation Intelligence Program",
"metrics":["25\nvalidated cases","13\ncountries/territories","680\nanimals/specimens in exact counts",">=21\nadditional in minimum counts"],
"foot":"Published 27 September 2026. The weekly core uses event/seizure date, not publication date alone.",
"scope":"Weekly core: only VALIDATED records with event/seizure dates between Monday 14 September 2026 00:00 and Sunday 20 September 2026 23:59:59, America/Manaus. Minimum and unknown quantities are excluded from the exact total.",
"sections":["1. Executive summary","2. Consolidated weekly core","3. Analytical reading of the week","4. Cases during the period","5. Limits of the evidence","6. Methodological note","7. Main sources"],
"executive":[
"The weekly core contains 25 VALIDATED cases across 13 countries or territories. Nineteen cases have exact counts totaling 680 animals or specimens. Two additional cases have minimum-only counts totaling at least 21 confirmed individuals, while four cases remain without a reliable numeric quantity.",
"Documented facts: the largest exact-count incident involved 206 birds on BR-222 in São Gonçalo do Amarante. Two incidents recorded 101 specimens each, in Singapore and Medan. In Samalayuca, 75 tarantulas and 20 turtles were seized; in Motu, 55 Indian flapshell turtles were intercepted.",
"Distinct logistics channels were documented - roads, bus transport, postal parcel, airport transfer baggage, air parcel and seaport - alongside exploitation or illegal hunting without a demonstrated trade route."
],
"analytical":[
["Observed patterns","Road transport is the primary mode in 8 of 25 records; 2 are classified as air and 1 as maritime. In 14 records, the primary mode was not reported. Birds and reptiles appear across a large share of the weekly core, while mammals are present in cases involving big cats, armadillos, pangolins and hedgehogs. A recurring transition from digital environments to physical seizure appears in WhatsApp offers in Angaco, online advertisements in Singapore and investigation of social-media sales in Mairiporã."],
["Analytical interpretation","The week's cases show a mosaic of supply and exploitation chains using commercial logistics, passenger baggage, road transport and digital channels. This reinforces the value of combining controls at transport nodes, digital intelligence and documentary traceability. The recurrence of different taxa through similar logistics channels suggests that risk indicators should prioritize behavior, route, documentation and concealment method, not only target species. These similarities do not, by themselves, demonstrate a common transnational network."],
["Limits of the evidence","Four cases lack reliable quantities and two have minimum-only counts, so 680 does not represent all animals involved. Sources for Motu differ by one day, but both dates fall inside the weekly window. In Singapore, the authority did not state that the six investigated people formed a single network. In Paulo Afonso, the Bahia-Pernambuco direction is not securely established. Biological origin, final buyer, complete route and network participation remain unknown in several records."]
],
"methodology":[
"Inclusion is based on event/seizure date between 14 and 20 September 2026 in the America/Manaus time zone; publication date is used only as supporting information.",
"Only VALIDATED records in the Observatory's canonical database are included. Cases whose event date could not be established inside the window are excluded from the weekly core.",
"Aggregated operations and joint releases were reconciled to prevent double counting. Hoedspruit and Gravelotte are distinct interventions, each involving one pangolin; jointly reported arrests and vehicles were not allocated between the two records.",
"Weight, monetary value, packages, wildlife parts and products were not automatically converted into numbers of animals. Volta Redonda remains uncounted because the source contains internally inconsistent figures."
],
"core_intro":"Chronological order by event/seizure date. Minimum or unknown quantities are excluded from the exact total of 680.",
"body_footer":"Created and managed by the Conservation Intelligence Program at WCS Brasil."
},
"es": {
"outfile":"Informe_Semanal_Observatorio_Global_14-20_Sep_2026_WCS.docx",
"title":"INFORME SEMANAL DE MONITOREO",
"subtitle":"Observatorio Global del Tráfico de Animales",
"period":"14-20 de septiembre de 2026 | Semana civil | America/Manaus",
"close":"WCS Brasil | Programa de Inteligencia para la Conservación",
"metrics":["25\ncasos validados","13\npaíses/territorios","680\nanimales/especímenes en conteos exactos",">=21\nadicionales en conteos mínimos"],
"foot":"Publicado el 27 de septiembre de 2026. El núcleo semanal usa la fecha de ocurrencia/decomiso, no solo la fecha de publicación.",
"scope":"Núcleo semanal: solo registros VALIDATED con ocurrencia/decomiso entre el lunes 14/09/2026 00:00 y el domingo 20/09/2026 23:59:59, America/Manaus. Las cantidades mínimas o desconocidas se excluyen del total exacto.",
"sections":["1. Resumen ejecutivo","2. Núcleo semanal consolidado","3. Lectura analítica de la semana","4. Casos del período","5. Límites de la evidencia","6. Nota metodológica","7. Fuentes principales"],
"executive":[
"El núcleo semanal reúne 25 casos VALIDATED en 13 países o territorios. Diecinueve casos tienen conteos exactos que suman 680 animales o especímenes. Dos casos adicionales presentan solo conteos mínimos, con al menos 21 individuos confirmados, mientras cuatro casos permanecen sin una cantidad numérica confiable.",
"Hechos documentados: el mayor incidente por conteo exacto involucró 206 aves en la BR-222, en São Gonçalo do Amarante. Dos incidentes registraron 101 especímenes cada uno, en Singapur y Medan. En Samalayuca se decomisaron 75 tarántulas y 20 tortugas; en Motu se interceptaron 55 tortugas de caparazón blando indias.",
"Se documentaron distintos canales logísticos - carreteras, autobús, encomienda postal, equipaje en conexión aeroportuaria, encomienda aérea y puerto marítimo - además de explotación o caza ilegal sin una ruta comercial demostrada."
],
"analytical":[
["Patrones observados","El transporte terrestre es el modo principal en 8 de 25 registros; 2 se clasifican como aéreos y 1 como marítimo. En 14 registros, el modo principal no fue informado. Aves y reptiles aparecen en gran parte del núcleo semanal, y los mamíferos están presentes en casos de grandes felinos, armadillos, pangolines y erizos. Se observa una transición recurrente de lo digital a la incautación física: WhatsApp en Angaco, anuncios en línea en Singapur e investigación de ventas por redes sociales en Mairiporã."],
["Interpretación analítica","Los casos de la semana muestran un mosaico de cadenas de suministro y explotación que utilizan logística comercial, equipaje de pasajeros, transporte por carretera y canales digitales. Esto refuerza el valor de combinar controles en nodos de transporte, inteligencia digital y trazabilidad documental. La recurrencia de taxones distintos en canales similares sugiere que los indicadores de riesgo deben priorizar comportamiento, ruta, documentación y método de ocultamiento, y no solo especies objetivo. Estas similitudes no demuestran, por sí mismas, una red transnacional común."],
["Límites de la evidencia","Cuatro casos carecen de cantidad confiable y dos presentan solo conteos mínimos; por eso 680 no representa a todos los animales involucrados. Las fuentes de Motu difieren en un día, pero ambas fechas están dentro de la ventana semanal. En Singapur, la autoridad no afirmó que las seis personas investigadas formaran una única red. En Paulo Afonso, la dirección Bahía-Pernambuco no se establece con seguridad. Origen biológico, comprador final, ruta completa y participación en redes siguen desconocidos en varios registros."]
],
"methodology":[
"La inclusión se basa en la fecha de ocurrencia/decomiso entre el 14 y el 20 de septiembre de 2026, en el huso America/Manaus; la fecha de publicación se usa solo como información auxiliar.",
"Solo se incluyen registros VALIDATED de la base canónica del Observatorio. Los casos cuya fecha de ocurrencia no pudo establecerse dentro de la ventana se excluyen del núcleo semanal.",
"Las operaciones agregadas y comunicados conjuntos fueron reconciliados para evitar doble conteo. Hoedspruit y Gravelotte son intervenciones distintas, cada una con un pangolín; las detenciones y vehículos divulgados conjuntamente no se repartieron entre los dos registros.",
"Peso, valor financiero, embalajes, partes y productos no se convirtieron automáticamente en número de animales. Volta Redonda permanece sin conteo por cifras internamente incompatibles en la fuente."
],
"core_intro":"Orden cronológico por fecha de ocurrencia/decomiso. Las cantidades mínimas o desconocidas se excluyen del total exacto de 680.",
"body_footer":"Creado y gestionado por el Programa de Inteligencia para la Conservación de WCS Brasil."
}
}

for lang, vals in updates.items():
    langs[lang].update(vals)

TEMP_LANGS.write_text(json.dumps(langs, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

src = BASE_SCRIPT.read_text(encoding="utf-8")
src = src.replace("import json\n", "import json\nimport re\n", 1)
src = src.replace("OUT=ROOT/'reports/2026/2026-09-08_2026-09-14'; OUT.mkdir(parents=True,exist_ok=True)",
                  "OUT=ROOT/'reports/2026/2026-09-14_2026-09-20'; OUT.mkdir(parents=True,exist_ok=True)")
src = src.replace("sources=json.loads((ROOT/'.github/data/backfill_sources.json').read_text(encoding='utf-8'))",
                  "sources=json.loads((ROOT/'.github/data/week_2026_09_14_20_sources.json').read_text(encoding='utf-8'))")
src = re.sub(r"cases=\[\]\nfor _n in \('a','b','c'\):\n    cases \+= json\.loads\(\(ROOT/f'\.github/data/backfill_cases_\{_n\}\.json'\)\.read_text\(encoding='utf-8'\)\)",
             "cases=json.loads((ROOT/'.github/data/week_2026_09_14_20_cases.json').read_text(encoding='utf-8'))", src)
src = src.replace("uncertain=json.loads((ROOT/'.github/data/backfill_uncertain.json').read_text(encoding='utf-8'))", "uncertain=[]")
src = src.replace("LANGS=json.loads((ROOT/'.github/data/backfill_LANGS.json').read_text(encoding='utf-8'))",
                  "LANGS=json.loads((ROOT/'.github/data/_week_2026_09_14_20_langs.json').read_text(encoding='utf-8'))")

old_body = """def body(doc,text,size=10,after=6,italic=False):
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(after); p.paragraph_format.line_spacing=1.08
    r=p.add_run(text); r.font.name='Aptos'; r.font.size=Pt(size); r.italic=italic
    return p
"""
new_body = """SCI_RE=re.compile(r'\\b(?:Chloropsis\\s+(?:moluccensis|cyanopogon|venusta)|C\\.\\s+(?:cyanopogon|venusta)|Carduelis\\s+carduelis|Panthera\\s+(?:leo|tigris)|Atelerix\\s+albiventris|Gekko\\s+gecko|Lissemys\\s+punctata|Brachypelma\\s+emilia|Kinosternon\\s+integrum|Pantherophis\\s+guttatus|Geochelone\\s+elegans)\\b')

def add_sci(p,text,size=10,italic=False):
    text=text.replace('*','')
    pos=0
    for m in SCI_RE.finditer(text):
        if m.start()>pos:
            r=p.add_run(text[pos:m.start()]); r.font.name='Aptos'; r.font.size=Pt(size); r.italic=italic
        r=p.add_run(m.group(0)); r.font.name='Aptos'; r.font.size=Pt(size); r.italic=True
        pos=m.end()
    if pos<len(text):
        r=p.add_run(text[pos:]); r.font.name='Aptos'; r.font.size=Pt(size); r.italic=italic

def body(doc,text,size=10,after=6,italic=False):
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(after); p.paragraph_format.line_spacing=1.08
    add_sci(p,text,size,italic)
    return p
"""
src = src.replace(old_body, new_body)
src = src.replace("p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER if i<2 else WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.space_after=Pt(0); r=p.add_run(x); r.font.name='Aptos'; r.font.size=Pt(6.8)",
                  "p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER if i<2 else WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.space_after=Pt(0); add_sci(p,x,6.8)")
src = re.sub(r"\n    heading\(doc,cfg\['sections'\]\[4\]\)\n    for u in uncertain:.*?\n    heading\(doc,cfg\['sections'\]\[5\]\)",
             "\n    heading(doc,cfg['sections'][5])", src, flags=re.S)

src = src.split("\nidx=ROOT/'reports.json'", 1)[0]
exec(compile(src, str(BASE_SCRIPT), "exec"), {"__name__":"__main__"})

record = {
"id": PERIOD_ID,
"period_start":"2026-09-14",
"period_end":"2026-09-20",
"year":2026,
"title_pt":"Relatório Semanal do Observatório Global de Tráfico de Animais - 14 a 20 de setembro de 2026",
"title_en":"Weekly Report of the Illegal Wildlife Trafficking Global Observatory - 14-20 September 2026",
"title_es":"Informe Semanal del Observatorio Global del Tráfico de Animales - 14-20 de septiembre de 2026",
"summary_pt":"25 casos validados em 13 países ou territórios; 680 animais ou espécimes em 19 contagens exatas, mais pelo menos 21 indivíduos em dois registros de contagem mínima. Quatro casos permanecem sem quantidade numérica confiável.",
"summary_en":"25 validated cases across 13 countries or territories; 680 animals or specimens in 19 exact-count cases, plus at least 21 individuals in two minimum-count records. Four cases remain without a reliable numeric quantity.",
"summary_es":"25 casos validados en 13 países o territorios; 680 animales o especímenes en 19 casos con conteo exacto, más al menos 21 individuos en dos registros de conteo mínimo. Cuatro casos permanecen sin una cantidad numérica confiable.",
"metrics":{"cases":25,"countries":13,"exact_animals":680,"exact_count_cases":19,"minimum_additional_animals":21,"minimum_count_cases":2,"unknown_quantity_cases":4},
"pdf_pt":f"reports/2026/{PERIOD_ID}/Relatorio_Semanal_Observatorio_Global_14-20_Set_2026_WCS.pdf",
"docx_pt":f"reports/2026/{PERIOD_ID}/Relatorio_Semanal_Observatorio_Global_14-20_Set_2026_WCS.docx",
"pdf_en":f"reports/2026/{PERIOD_ID}/Weekly_Report_Global_Observatory_14-20_Sep_2026_WCS.pdf",
"docx_en":f"reports/2026/{PERIOD_ID}/Weekly_Report_Global_Observatory_14-20_Sep_2026_WCS.docx",
"pdf_es":f"reports/2026/{PERIOD_ID}/Informe_Semanal_Observatorio_Global_14-20_Sep_2026_WCS.pdf",
"docx_es":f"reports/2026/{PERIOD_ID}/Informe_Semanal_Observatorio_Global_14-20_Sep_2026_WCS.docx",
"published_at":"2026-09-27"
}
index_path=ROOT/"reports.json"
index=json.loads(index_path.read_text(encoding="utf-8"))
index=[r for r in index if r.get("id")!=PERIOD_ID]
index.insert(0,record)
index_path.write_text(json.dumps(index,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
TEMP_LANGS.unlink(missing_ok=True)
