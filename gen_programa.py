import subprocess
import pandas as pd
from pylatex import Document, Package, Command, PageStyle, Head, Foot, NewPage, NewLine,\
    TextColor, MiniPage, StandAloneGraphic, simple_page_number,\
    TikZ, TikZNode, TikZOptions, TikZCoordinate,\
    VerticalSpace, HorizontalSpace,\
    LongTabularx, Tabularx,\
    config
from pylatex.base_classes import Arguments
from pylatex.utils import NoEscape, bold
import funciones as fun

cursos = pd.read_csv("cursos/cursos_malla.csv")
cursos = cursos.fillna("")
detall = pd.read_csv("cursos/cursos_detalles.csv")
progra = pd.read_csv("cursos/cursos_programas.csv")
descri = pd.read_csv("cursos/cursos_descri.csv")
atribu = pd.read_csv("cursos/cursos_atributos.csv")
objeti = pd.read_csv("cursos/cursos_obj.csv")
conten = pd.read_csv("cursos/cursos_conten.csv")
metodo = pd.read_csv("cursos/cursos_metodo.csv")
metdes = pd.read_csv("cursos/descri_metodo.csv")
evalua = pd.read_csv("cursos/cursos_evalua.csv")
evades = pd.read_csv("cursos/descri_evalua.csv")
evatip = pd.read_csv("cursos/tipos_evalua.csv")
bibtex = pd.read_csv("cursos/cursos_bibtex.csv")
profes = pd.read_csv("cursos/cursos_profes.csv")
datpro = pd.read_csv("profes/profes_datos.csv")
grapro = pd.read_csv("profes/profes_grados.csv")

tipCursoDic = {
    0: "Teórico",
    1: "Práctico",
    2: "Teórico - Práctico"
}

eleCursoDic = {
    0: "Obligatorio",
    1: "Electivo"
}

tipAsistDic = {
    0: "Libre",
    1: "Obligatoria" 
}

sinoDic = {
    0: "No",
    1: "Sí" 
}


def textcolor(size,vspace,color,bold,text,hspace="0",par=True):
    dump = NoEscape(r"")
    if par==True:
        dump = NoEscape(r"\par")
    if hspace!="0":
        dump += NoEscape(HorizontalSpace(hspace,star=True).dumps())
    dump += NoEscape(Command("fontsize",arguments=Arguments(size,vspace)).dumps())
    dump += NoEscape(Command("selectfont").dumps()) + NoEscape(" ")
    if bold==True:
        dump += NoEscape(Command("textbf", NoEscape(Command("textcolor",arguments=Arguments(color,text)).dumps())).dumps())
    else:
        dump += NoEscape(Command("textcolor",arguments=Arguments(color,text)).dumps())
    return dump

def fontselect(size,vspace):
    dump = NoEscape(r"")
    dump += NoEscape(Command("fontsize",arguments=Arguments(size,vspace)).dumps())
    dump += NoEscape(Command("selectfont").dumps()) + NoEscape(" ")
    return dump

def generar_programa(id):
    listProf = profes[profes.id == id].profesores.str.split(';').item()
    codCurso = cursos[cursos.id == id].codigo.item()
    nomEscue = "Escuela de Física"
    lisProgr = progra[progra.id == id].drop('id',axis=1)
    numProgr = len(lisProgr.programa)
    counter = 0
    if numProgr > 1:
        strProgr = "Carreras de: "
    else:
        strProgr = "Carrera de "
    for programa in lisProgr.programa:
        counter += 1
        strProgr += programa
        if counter < numProgr:
            if numProgr == 2:
                strProgr += " y "
            else:
                strProgr += ", "  
    nomCurso = cursos[cursos.id == id].nombre.item()
    print(f'Curso: {nomCurso}')
    tipo = detall[detall.id == id].tipo.item()
    tipCurso = tipCursoDic.get(detall[detall.id == id].tipo.item())
    eleCurso = eleCursoDic.get(detall[detall.id == id].electivo.item())
    numCredi = cursos[cursos.id == id].creditos.item()
    horClass = cursos[cursos.id == id].horasTeoria.item() + cursos[cursos.id == id].horasPractica.item()
    horExtra = (numCredi * 3) - horClass
    semCurso = cursos[cursos.id == id].semestre.item()
    ubiPlane = ""
    counter = 0
    for programa in lisProgr.programa:
        counter +=1
        semestre = lisProgr[lisProgr['programa'] == programa].semestre.item()
        if semestre <= 10:
            ubiPlane += "Curso de " + fun.number_to_ordinals(str(int(semestre))) + " semestre en " + programa
        elif semestre > 10:
            ubiPlane += "Curso electivo en " + programa
        if counter > 1:
            ubiPlane += " "
    tipAsist = tipAsistDic.get(detall[detall.id == id].asistencia.item())
    posSufic = sinoDic.get(detall[detall.id == id].suficiencia.item())
    posRecon = sinoDic.get(detall[detall.id == id].reconocimiento.item())
    aprCurso = detall[detall.id == id].aprobacion.str.split(';').explode().reset_index(drop=True)
    aprCurso = aprCurso[0] + "/" + aprCurso[1] + "/" + aprCurso[2] + " en sesión de Consejo de Escuela " + aprCurso[3]
    desGener = NoEscape(descri[descri.id == id].descripcion.item().replace('\n', r'\newline\newline '))
    desGener += NoEscape(r"El curso busca desarrollar los siguientes atributos de egreso: \newline")
    lisAtrib = atribu[atribu.id == id].reset_index(drop=True)
    atrTabla = NoEscape(r" \begin{minipage}{\linewidth} ")
    atrTabla += NoEscape(r" \centering ") 
    atrTabla += NoEscape(r" \begin{tabular}{ p{5cm}  p{2cm} } ")
    atrTabla += NoEscape(r" \toprule ") 
    for consecutivo, atribs in lisAtrib.iterrows():
        atrTabla += NoEscape(f" {atribs.atributo} & {atribs.nivel}") + NoEscape(r" \\ ")
        if consecutivo == len(lisAtrib)-1:
            atrTabla += NoEscape(r" \bottomrule ")
        else: 
            atrTabla += NoEscape(r" \midrule ")
    atrTabla += NoEscape(r" \end{tabular} \end{minipage}")
    lisObjet = objeti[objeti.id == id].reset_index(drop=True).objetivo
    for consecutivo, objetivo in lisObjet.items():
        if consecutivo == 0:
            objGener = NoEscape(objetivo)
            objEspec = NoEscape(r"\begin{itemize}")
        else:
            objEspec += NoEscape(r"\item ") + NoEscape(objetivo) + NoEscape(r".")
    objEspec += NoEscape(r"\end{itemize}")
    objCurso = NoEscape(r"Al final del curso la persona estudiante será capaz de:") 
    objCurso += NoEscape(r"\newline\newline ")
    objCurso += NoEscape(Command("textbf", "Objetivo general").dumps())
    objCurso += NoEscape(r"\begin{itemize}\item ")
    objCurso += objGener + NoEscape(r".")
    objCurso += NoEscape(r"\end{itemize} \vspace{2mm}")
    objCurso += NoEscape(Command("textbf", "Objetivos específicos").dumps())
    objCurso += objEspec
    if tipo == 1:
        if codCurso not in ["EE9001","EE1102"]:
            conDescr = "En el curso se desarrollaran los siguientes laboratorios:"
        else:
            conDescr = "En el curso se desarrollarán los siguientes temas:"
    else:
        conDescr = "En el curso se desarrollaran los siguientes temas:"
    conCurso = NoEscape(r"\par \setlength{\leftskip}{4cm} ")
    conCurso += NoEscape(r"\begin{easylist} \ListProperties(Progressive*=3ex)")
    conCurso += NoEscape(conten[conten.id == id].contenidos.item())
    conCurso += NoEscape(r"\end{easylist} ")
    conCurso += NoEscape(r"\setlength{\leftskip}{0cm} ")
    lisMetod = metodo[metodo.id == id].reset_index(drop=True).metodologia
    for consecutivo, metodos in lisMetod.items():
        if consecutivo == 0:
            metGener = NoEscape(metdes[metdes["tipo"]==tipo].descripcion.item())
            metEspec = NoEscape(r"\begin{itemize}")
            metEspec += NoEscape(r"\item ") + NoEscape(metodos)
        else:
            metEspec += NoEscape(r"\item ") + NoEscape(metodos)
    metEspec += NoEscape(r"\end{itemize}")
    metCurso = metGener
    metCurso += NoEscape(r"\newline\newline ")
    metCurso += NoEscape(Command("textbf", "Las personas estudiantes desarrollarán:").dumps() + r" \newline")
    metCurso += metEspec
    metCurso += NoEscape(r"\vspace*{2mm}")
    metCurso += NoEscape(f"Este enfoque metodológico permitirá a la persona estudiante {objGener[0].lower() + objGener[1:]}.")
    metCurso += NoEscape(r"\vspace*{2mm} \newline  ")
    metCurso += NoEscape(r"Si un estudiante requiere apoyos educativos, podrá solicitarlos a través del Departamento de Orientación y Psicología. \newline ")
    evaCurso = NoEscape(r"La evaluación se distribuye en los siguientes rubros:")
    evaCurso += NoEscape(r" \newline ")
    tipEvalu = evalua[evalua.id == id].tipoEval.item()
    lisEvalu = evatip[(evatip.tipo == tipo) & (evatip.tipoEval == tipEvalu)].reset_index(drop=True)
    for consecutivo, evaluas in lisEvalu.iterrows():
        if consecutivo == 0:
            descriEval = NoEscape(r"\begin{itemize} ")  
        descriEval += NoEscape(r"\item ") + NoEscape(f"{evaluas.evaluacion}: {evades[evades["evaluacion"]==evaluas.evaluacion].descripcion.item()}")  
    descriEval += NoEscape(r"\end{itemize}") 
    evaCurso += descriEval
    evaTabla = NoEscape(r" \begin{minipage}{\linewidth} ")
    evaTabla += NoEscape(r" \centering ") 
    evaTabla += NoEscape(r" \begin{tabular}{ p{4.5cm}  p{1.5cm} } ")
    evaTabla += NoEscape(r" \toprule ") 
    total = 0
    for consecutivo, evaluas in lisEvalu.iterrows():
        evaTabla += NoEscape(f" {evaluas.evaluacion} ({evaluas.cantidad}) & {evaluas.porcentaje} \\%") + NoEscape(r" \\ ")
        evaTabla += NoEscape(r" \midrule ")
        total += evaluas.porcentaje
        if consecutivo == len(lisEvalu)-1:
            evaTabla += NoEscape(f"Total & {total} \\%") + NoEscape(r" \\ ")
            evaTabla += NoEscape(r" \bottomrule ")
    evaTabla += NoEscape(r" \end{tabular} \end{minipage}")
    evaRepo = NoEscape(r"De conformidad con el artículo 78 del Reglamento del Régimen Enseñanza-Aprendizaje del Instituto Tecnológico de Costa Rica y sus Reformas, en este curso la persona estudiante ")
    if tipCurso != "Teórico":
        evaRepo += NoEscape(Command("textbf", "no").dumps())
    evaRepo += NoEscape(r" tiene derecho a presentar un examen de reposición")
    if tipCurso == "Teórico":
        evaRepo += NoEscape(r" si su nota luego de redondeo es 60 o 65.")
    else:
        evaRepo += NoEscape(r".")
    bibCurso = NoEscape(r'\nocite{' + ('} '+r'\nocite{').join(bibtex[bibtex.id == id].bibtex.item().split(';')) + '} ')
    bibPrint = NoEscape(r'\vspace*{-8mm}\printbibliography[heading=none]')
    dataProf = datpro[datpro.codigo.isin(listProf)]
    proImpar = NoEscape(r"El curso será impartido por:")
    proCurso = NoEscape(r'\vspace*{-4mm}\begin{textoMargen}')
    for consecutivo, profe in dataProf.iterrows():
        print(f'Profesor: {profe.nombre}')
        match profe.titulo:
            case "M.Sc." | "Lic." | "Máster" | "Dr.-Ing." | "Mag.":
                proCurso += NoEscape(Command("textbf", f"{profe.titulo} {profe.nombre}").dumps())
            case "Ph.D.":
                proCurso += NoEscape(Command("textbf", f"{profe.nombre}, {profe.titulo}").dumps())
        proCurso += NoEscape(r" \newline ")
        gradProf = grapro[grapro.codigo == profe.codigo]
        for consecutivo, grado in gradProf.iterrows():
            proCurso += NoEscape(Command("textbf", f"{grado.grado} en {grado.campo}, {grado.institucion}, {grado.pais}").dumps())
            proCurso += NoEscape(r" \newline \newline ") 
        proCurso += NoEscape(Command("emph", "Correo:").dumps())               
        proCurso += NoEscape(f" {profe.correo}")
        proCurso += NoEscape(Command("emph", "  Teléfono:").dumps())     
        proCurso += NoEscape(f" {int(profe.telefono)}")
        proCurso += NoEscape(r" \vspace*{1mm} \newline ")
        proCurso += NoEscape(Command("emph", "  Oficina:").dumps())    
        proCurso += NoEscape(f" {int(profe.oficina)}")
        proCurso += NoEscape(Command("emph", "  Escuela:").dumps())  
        proCurso += NoEscape(f" {profe.escuela}")
        proCurso += NoEscape(Command("emph", "  Sede:").dumps())  
        proCurso += NoEscape(f" {profe.sede}")
        proCurso += NoEscape(r" \vspace*{4mm} \newline ")             
    proCurso += NoEscape(r"\end{textoMargen}")
    #Config
    config.active = config.Version1(row_heigth=1.5)
    #Geometry
    geometry_options = { 
        "left": "22.5mm",
        "right": "16.1mm",
        "top": "48mm",
        "bottom": "25mm",
        "headheight": "12.5mm",
        "footskip": "12.5mm"
    }
    #Document options
    doc = Document(documentclass="article", \
                   fontenc=None, \
                   inputenc=None, \
                   lmodern=False, \
                   textcomp=False, \
                   page_numbers=True, \
                   indent=False, \
                   document_options=["letterpaper"],
                   geometry_options=geometry_options)
    #Packages
    doc.packages.append(Package(name="fontspec", options=None))
    doc.packages.append(Package(name="babel", options=['spanish','activeacute']))
    doc.packages.append(Package(name="anyfontsize"))
    doc.packages.append(Package(name="fancyhdr"))
    doc.packages.append(Package(name="csquotes"))
    doc.packages.append(Package(name="easylist", options=['ampersand']))
    doc.packages.append(Package(name="biblatex", options=['style=ieee','backend=biber']))
    doc.packages.append(Package(name="tcolorbox",options=['skins','breakable']))
    doc.packages.append(Package(name="booktabs"))
    #Package options
    doc.preamble.append(Command('setmainfont','Arial'))
    doc.preamble.append(Command('addbibresource', '../bibIF.bib'))
    doc.preamble.append(NoEscape(r'\renewcommand*{\bibfont}{\fontsize{10}{14}\selectfont}'))
    doc.preamble.append(NoEscape(r'''
\defbibenvironment{bibliography}
    {\list
    {\printfield[labelnumberwidth]{labelnumber}}
    {\setlength{\leftmargin}{4cm}
    \setlength{\rightmargin}{1.1cm}
    \setlength{\itemindent}{0pt}
    \setlength{\itemsep}{\bibitemsep}
    \setlength{\parsep}{\bibparsep}}}
    {\endlist}
{\item}
'''))
    doc.preamble.append(NoEscape(r'''
\newenvironment{textoMargen}
    {%
    \begin{list}{}{%
        \setlength{\leftmargin}{3.6cm}%
        \setlength{\rightmargin}{1.1cm}%
    }%
    \item[]%
  }
  {%
    \end{list}%
  }
'''))
    doc.add_color('gris','rgb','0.27,0.27,0.27') #70,70,70
    doc.add_color('parte','rgb','0.02,0.204,0.404') #5,52,103
    doc.add_color('azulsuaveTEC','rgb','0.02,0.455,0.773') #5,116,197
    doc.add_color('fila','rgb','0.929,0.929,0.929') #237,237,237
    doc.add_color('linea','rgb','0.749,0.749,0.749') #191,191,191

    headerfooter = PageStyle("headfoot")

    #Left header
    with headerfooter.create(Head("L")) as header_left:
        with header_left.create(MiniPage(width=r"0.5\textwidth",align="l")) as logobox:
            logobox.append(StandAloneGraphic(image_options="width=62.5mm", filename='../figuras/Logo.png'))
    #Right foot
    with headerfooter.create(Foot("R")) as footer_right:
        footer_right.append(TextColor("black", NoEscape(r"Página \thepage \hspace{1pt} de \pageref*{LastPage}")))        
    #Add header and footer 
    doc.preamble.append(headerfooter)
    doc.change_page_style("empty")
    #Set logo in first page
    with doc.create(TikZ(
            options=TikZOptions
                (    
                "overlay",
                "remember picture"
                )
        )) as logo:
        logo.append(TikZNode(\
            options=TikZOptions
                (
                "inner sep = 0mm",
                "outer sep = 0mm",
                "anchor = north west",
                "xshift = -23mm",
                "yshift = 22mm"
                ),
            text=StandAloneGraphic(image_options="width=21cm", filename='../figuras/Logo_portada.png').dumps(),\
            at=TikZCoordinate(0,0)
        ))
    doc.append(VerticalSpace("100mm", star=True))
    doc.append(textcolor
            (   
            size="14",
            vspace="0",
            color="black",
            bold=False,
            text=f"Programa del curso {str(codCurso)[:2]}-{str(codCurso)[2:]}"
            ))
    doc.append(textcolor
            (
            size="18",
            vspace="25",
            color="black",
            bold=True,
            text=f"{nomCurso}" 
            ))
    doc.append(VerticalSpace("15mm", star=True))
    doc.append(NewLine())
    with doc.create(Tabularx(table_spec=r"m{0.02\textwidth}m{0.98\textwidth}")) as table:
            table.add_row(["", textcolor
            (   
            par=False,
            hspace="0mm",
            size="12",
            vspace="0",
            color="gris",
            bold=True,
            text=f"{nomEscue}"
            )])
            table.append(NoEscape('[-12pt]'))
            table.add_row(["", textcolor
            (   
            par=False,
            hspace="0mm",
            size="12",
            vspace="0",
            color="gris",
            bold=True,
            text=f"{strProgr}" 
            )])
    doc.append(NewPage())
    doc.change_document_style("headfoot")
    doc.append(textcolor
            (   
            size="14",
            vspace="0",
            color="parte",
            bold=True,
            text="I parte: Aspectos relativos al plan de estudios"
            ))
    doc.append(textcolor
            (   
            hspace="2mm",
            size="12",
            vspace="14",
            color="parte",
            bold=True,
            text="1. Datos generales"
            ))
    doc.append(VerticalSpace("3mm", star=True))
    doc.append(NewLine())
    doc.append(fontselect
            (
            size="10",
            vspace="12"      
            ))
    with doc.create(Tabularx(table_spec=r"p{6cm}p{10cm}")) as table:
            table.add_row([bold("Nombre del curso:"), f"{nomCurso}"])
            table.append(NoEscape('[10pt]'))
            table.add_row([bold("Código:"), f"{str(codCurso)}"])
            table.append(NoEscape('[10pt]'))
            table.add_row([bold("Tipo de curso:"), f"{tipCurso}"])
            table.append(NoEscape('[10pt]'))
            table.add_row([bold("Obligatorio o electivo:"), f"{eleCurso}"])
            table.append(NoEscape('[10pt]'))
            table.add_row([bold("Nº de créditos:"), f"{numCredi}"])
            table.append(NoEscape('[10pt]'))
            table.add_row([bold("Nº horas de clase por semana:"), f"{horClass}"])
            table.append(NoEscape('[10pt]'))
            table.add_row([bold("Nº horas extraclase por semana:"), f"{horExtra}"])
            table.append(NoEscape('[10pt]'))
            table.add_row([bold("Ubicación en el plan de estudios:"), NoEscape(f"{ubiPlane}")])
            table.append(NoEscape('[10pt]'))
            lisRequi = cursos[cursos.id == id].requisitos.str.split(";").explode().reset_index(drop=True)
            for consecutivo, requisito in lisRequi.items():
                if consecutivo == 0 and requisito == "":
                    table.add_row([bold("Requisitos:"), "Ninguno"])
                elif consecutivo == 0:   
                    table.add_row([bold("Requisitos:"), f"{cursos[cursos.id == requisito].codigo.item()} \
                                    {cursos[cursos.id == requisito].nombre.item()}"])   
                else: 
                    table.add_row(["", f"{cursos[cursos.id == requisito].codigo.item()} \
                                    {cursos[cursos.id == requisito].nombre.item()}"])  
                table.append(NoEscape('[10pt]'))  
            lisCorre = cursos[cursos.id == id].correquisitos.str.split(';').explode().reset_index(drop=True)
            for consecutivo, correquisito in lisCorre.items():
                if consecutivo == 0 and correquisito == "":
                    table.add_row([bold("Correquisitos:"), "Ninguno"])
                elif consecutivo == 0:
                    table.add_row([bold("Correquisitos:"), f"{cursos[cursos.id == correquisito].codigo.item()}\
                                    {cursos[cursos.id == correquisito].nombre.item()}"])   
                else: 
                    table.add_row(["", f"{cursos[cursos.id == correquisito].codigo.item()}\
                                    {cursos[cursos.id == correquisito].nombre.item()}"])  
                table.append(NoEscape('[10pt]')) 
            lisEsreq = cursos[cursos.id == id].esrequisito.str.split(';').explode().reset_index(drop=True)
            for consecutivo, esrequisito in lisEsreq.items():
                if consecutivo == 0 and esrequisito == "":
                    table.add_row([bold("El curso es requisito de:"), "Ninguno"])
                elif consecutivo == 0:
                    table.add_row([bold("El curso es requisito de:"), f"{cursos[cursos.id == esrequisito].codigo.item()}\
                                    {cursos[cursos.id == esrequisito].nombre.item()}"])   
                else: 
                    table.add_row(["", f"{cursos[cursos.id == esrequisito].codigo.item()}\
                                    {cursos[cursos.id == esrequisito].nombre.item()}"])  
                table.append(NoEscape('[10pt]'))        

    #             essRequi = NoEscape("")
    # lisEsreq = cursos[cursos.id == id].esrequisito.item()
    # texRequi = ""
    # if str(lisEsreq) != "nan":
    #     lisEsreq = cursos[cursos.id == id].esrequisito.str.split(';').explode().reset_index(drop=True)
    #     for esrequisito in lisEsreq:      
    #         texRequi += cursos[cursos.id == esrequisito].codigo.item()[:2] + "-" + cursos[cursos.id == esrequisito].codigo.item()[2:]
    #         texRequi += " "
    #         texRequi += cursos[cursos.id == esrequisito].nombre.item()
    #     essRequi = NoEscape(texRequi)
    # else:
    #     essRequi += NoEscape("Ninguno")    
            table.add_row([bold("Asistencia:"), f"{tipAsist}"])
            table.append(NoEscape('[10pt]'))
            table.add_row([bold("Suficiencia:"), f"{posSufic}"])
            table.append(NoEscape('[10pt]'))
            table.add_row([bold("Posibilidad de reconocimiento:"), f"{posRecon}"])
            table.append(NoEscape('[10pt]'))
            table.add_row([bold("Aprobación y actualización del programa:"), f"{aprCurso}"])
            table.append(NoEscape('[10pt]'))
    doc.append(NewPage())
    with doc.create(Tabularx(table_spec=r"p{3cm}p{13cm}")) as table:
            table.add_row([textcolor
            (   
            size="12",
            vspace="14",
            color="parte",
            bold=True,
            text="2. Descripción general"
            )
            ,desGener])  
    doc.append(VerticalSpace("2mm", star=True))  
    doc.append(NewLine())
    doc.append(atrTabla)
    doc.append(VerticalSpace("4mm", star=True))  
    doc.append(NewLine())
    with doc.create(Tabularx(table_spec=r"p{3cm}p{13cm}")) as table:
            table.add_row([textcolor
            (   
            size="12",
            vspace="14",
            color="parte",
            bold=True,
            text="3. Objetivos"
            )
            ,objCurso])
    doc.append(VerticalSpace("4mm", star=True)) 
    doc.append(NewLine())
    with doc.create(Tabularx(table_spec=r"p{3cm}p{13cm}")) as table:
        table.add_row([textcolor
        (   
        size="12",
        vspace="14",
        color="parte",
        bold=True,
        text="4. Contenidos"
        )
        ,conDescr])
    doc.append(NewLine())
    doc.append(conCurso)
    doc.append(textcolor
        (   
        size="14",
        vspace="0",
        color="parte",
        bold=True,
        text="II parte: Aspectos operativos"
        ))
    doc.append(VerticalSpace("4mm", star=True))  
    doc.append(NewLine())
    doc.append(fontselect
        (
        size="10",
        vspace="12"      
        ))
    with doc.create(Tabularx(table_spec=r"p{3cm}p{13cm}")) as table:
        table.add_row([textcolor
        (   
        size="12",
        vspace="14",
        color="parte",
        bold=True,
        text="5. Metodología"
        )
        ,metCurso])
    doc.append(VerticalSpace("2mm", star=True))  
    doc.append(NewLine())
    with doc.create(Tabularx(table_spec=r"p{3cm}p{13cm}")) as table:
        table.add_row([textcolor
        (   
        size="12",
        vspace="14",
        color="parte",
        bold=True,
        text="6. Evaluación"
        )
        ,evaCurso])
    doc.append(VerticalSpace("2mm", star=True))  
    doc.append(NewLine())
    doc.append(evaTabla)
    doc.append(VerticalSpace("2mm", star=True))  
    doc.append(NewLine())
    with doc.create(Tabularx(table_spec=r"p{3cm}p{13cm}")) as table:
        table.add_row([""
        ,evaRepo])
    doc.append(VerticalSpace("4mm", star=True))  
    doc.append(NewLine()) #antes era newline
    with doc.create(Tabularx(table_spec=r"p{3cm}p{13cm}")) as table:
        table.add_row([textcolor
        (   
        size="12",
        vspace="14",
        color="parte",
        bold=True,
        text="7. Bibliografía"
        )
        ,bibCurso]) 
    doc.append(bibPrint)
    # doc.append(VerticalSpace("2mm", star=True))  
    # doc.append(NewLine())
    with doc.create(Tabularx(table_spec=r"p{3cm}p{13cm}")) as table:
        table.add_row([textcolor
        (   
        size="12",
        vspace="14",
        color="parte",
        bold=True,
        text="8. Persona docente"
        )
        ,proImpar])
    doc.append(proCurso)
    doc.generate_pdf(f"./programas/{codCurso}", clean=False, clean_tex=False, compiler='lualatex')
    subprocess.run(["biber", f"C:\\Repositories\\IF\\programas\\{codCurso}"])
    doc.generate_pdf(f"./programas/{codCurso}", clean=False, clean_tex=False, compiler='lualatex')
    doc.generate_pdf(f"./programas/{codCurso}", clean=False, clean_tex=False, compiler='lualatex') 
    subprocess.run(f'move "C:\\Repositories\\IF\\programas\\{codCurso}.pdf" "C:\\Repositories\\IF\\programas\\IFI-{id[3:7]}-{codCurso}-{nomCurso}.pdf"', shell=True, check=True)

generar_programa("IFI0402") #Intrumentación I


subprocess.run(["del", f"C:\\Repositories\\IF\\programas\\*.tex"], shell=True, check=True)
subprocess.run(["del", f"C:\\Repositories\\IF\\programas\\*.aux"], shell=True, check=True)
subprocess.run(["del", f"C:\\Repositories\\IF\\programas\\*.bbl"], shell=True, check=True)
subprocess.run(["del", f"C:\\Repositories\\IF\\programas\\*.bcf"], shell=True, check=True)
subprocess.run(["del", f"C:\\Repositories\\IF\\programas\\*.blg"], shell=True, check=True)
subprocess.run(["del", f"C:\\Repositories\\IF\\programas\\*.log"], shell=True, check=True)
subprocess.run(["del", f"C:\\Repositories\\IF\\programas\\*.run.xml"], shell=True, check=True)