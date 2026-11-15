"""
Servicio generador de Informes Técnicos Periciales de Inspección SST en PDF con ReportLab.
Estructurado rigurosamente según el marco legal y normas técnicas colombianas:
- NTC 1461: Colores y señales de seguridad (código de colores, formas y porcentajes de área).
- NTC-ISO 3864-1:2011 y NTC-ISO 3864-3:2012: Principios de diseño para señales de seguridad.
- NTC-ISO 7010:2021: Símbolos gráficos registrados.
- Resolución 2400 de 1979 (Estatuto de Seguridad Industrial: pasillos libres >= 40cm, demarcación amarilla 10cm, barandas 60cm a 45°).
- NSR-10: Título J (Protección contra incendios, NTC 2885 / NTC 1669) y Título K (Medios y señalización de evacuación).
- Decreto 1072 de 2015: Art. 2.2.4.6.25 (Plan de Prevención, Preparación y Respuesta ante Emergencias).
- ISO 16069: Sistemas de señalización fotoluminiscente de evacuación (señales altas, intermedias y a nivel de zócalo).
- Catálogos técnicos oficiales (CCS y Covadonga): Fórmula S >= d² / 2000, distancias vs. tamaños y fotoluminiscencia Clase A/B.
- Estándar de citación y registro fotográfico APA 7ma Edición.
"""
import io
import os
import base64
from pathlib import Path
from typing import Dict, Any, List, Optional
from datetime import datetime
from PIL import Image as PILImage

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

from app.config import UPLOADS_DIR, BASE_DIR
from app.models.catalogo_sst import CATEGORIAS_SST, CATALOGO_SENALES


class NumberedCanvas(canvas.Canvas):
    """Canvas dinámico de ReportLab que calcula y numera las páginas totales (Página X de Y)."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("Helvetica-Bold", 7.5)
        self.setFillColor(colors.HexColor("#475569"))
        
        # Encabezado institucional en páginas 2 en adelante
        if self._pageNumber > 1:
            self.drawString(36, 762, "Politécnico Grancolombiano | Auditoría Pericial SST: Condiciones, Señalización y Demarcación")
            self.drawRightString(576, 762, "NTC 1461 / NSR-10 / Res. 2400 / Dec. 1072")
            self.setStrokeColor(colors.HexColor("#cbd5e1"))
            self.setLineWidth(0.6)
            self.line(36, 756, 576, 756)

        # Pie de página institucional en todas las páginas
        self.setStrokeColor(colors.HexColor("#cbd5e1"))
        self.setLineWidth(0.6)
        self.line(36, 42, 576, 42)
        
        self.setFont("Helvetica", 7.5)
        self.setFillColor(colors.HexColor("#64748b"))
        texto_pie = f"Página {self._pageNumber} de {page_count} · Dictamen Oficial de Auditoría Técnica y Conformidad Legal SST"
        self.drawString(36, 28, texto_pie)
        self.drawRightString(576, 28, f"Sistema SST-IA · {datetime.now().strftime('%d/%m/%Y')}")
        self.restoreState()


class PDFService:
    def __init__(self):
        self.styles = getSampleStyleSheet()
        self._crear_estilos()

    def _crear_estilos(self):
        self.estilo_titulo = ParagraphStyle(
            'TituloSST',
            parent=self.styles['Heading1'],
            fontName='Helvetica-Bold',
            fontSize=14,
            leading=18,
            textColor=colors.HexColor('#0f172a'),
            spaceAfter=4
        )
        self.estilo_subtitulo_meta = ParagraphStyle(
            'SubtituloMetaSST',
            parent=self.styles['Normal'],
            fontName='Helvetica',
            fontSize=8,
            leading=11.5,
            textColor=colors.HexColor('#334155'),
            spaceAfter=6
        )
        self.estilo_seccion_num = ParagraphStyle(
            'SeccionNumSST',
            parent=self.styles['Heading2'],
            fontName='Helvetica-Bold',
            fontSize=10.5,
            leading=13.5,
            textColor=colors.HexColor('#1e3a8a'),
            spaceBefore=9,
            spaceAfter=4
        )
        self.estilo_parrafo = ParagraphStyle(
            'ParrafoSST',
            parent=self.styles['Normal'],
            fontName='Helvetica',
            fontSize=7.8,
            leading=11,
            textColor=colors.HexColor('#1e293b'),
            alignment=4, # Justificado
            spaceAfter=4
        )
        self.estilo_norma_item = ParagraphStyle(
            'NormaItemSST',
            parent=self.styles['Normal'],
            fontName='Helvetica',
            fontSize=7.5,
            leading=10.5,
            textColor=colors.HexColor('#1e293b'),
            spaceAfter=3
        )
        self.estilo_celda = ParagraphStyle(
            'CeldaSST',
            parent=self.styles['Normal'],
            fontName='Helvetica',
            fontSize=6.8,
            leading=9,
            textColor=colors.HexColor('#1e293b')
        )
        self.estilo_celda_bold = ParagraphStyle(
            'CeldaBoldSST',
            parent=self.styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=6.8,
            leading=9,
            textColor=colors.HexColor('#0f172a')
        )
        self.estilo_celda_center = ParagraphStyle(
            'CeldaCenterSST',
            parent=self.styles['Normal'],
            fontName='Helvetica',
            fontSize=6.8,
            leading=9,
            alignment=1,
            textColor=colors.HexColor('#1e293b')
        )
        self.estilo_celda_th = ParagraphStyle(
            'CeldaThSST',
            parent=self.styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=7.2,
            leading=9.5,
            textColor=colors.HexColor('#ffffff'),
            alignment=1
        )
        # Estilos estándar APA 7ma Edición para figuras
        self.estilo_apa_figura = ParagraphStyle(
            'ApaFiguraSST',
            parent=self.styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=8.5,
            leading=11,
            textColor=colors.HexColor('#0f172a'),
            spaceAfter=1
        )
        self.estilo_apa_titulo = ParagraphStyle(
            'ApaTituloSST',
            parent=self.styles['Normal'],
            fontName='Helvetica-Oblique',
            fontSize=7.8,
            leading=10,
            textColor=colors.HexColor('#334155'),
            spaceAfter=4
        )
        self.estilo_apa_nota = ParagraphStyle(
            'ApaNotaSST',
            parent=self.styles['Normal'],
            fontName='Helvetica',
            fontSize=7,
            leading=9.5,
            textColor=colors.HexColor('#475569'),
            spaceBefore=3,
            spaceAfter=8
        )

    def _cargar_imagen_flowable(self, ev: Dict[str, Any], max_w=4.8 * inch, max_h=2.6 * inch) -> Optional[RLImage]:
        """Extrae la imagen real de la evidencia (Base64, archivo o static upload) y crea un RLImage escalado."""
        b64_candidato = None

        # 1. Buscar en hallazgos imagen anotada por OpenCV
        for h in ev.get("hallazgos", []):
            if h.get("foto_anotada_base64"):
                b64_candidato = h.get("foto_anotada_base64")
                break

        # 2. Buscar en la evidencia directa foto_base64
        if not b64_candidato and ev.get("foto_base64"):
            b64_candidato = ev.get("foto_base64")

        if b64_candidato:
            try:
                clean = b64_candidato.split(",", 1)[-1] if "," in b64_candidato else b64_candidato
                img_data = base64.b64decode(clean)
                pil_img = PILImage.open(io.BytesIO(img_data)).convert("RGB")
                w, h = pil_img.size
                ratio = min(max_w / w, max_h / h, 1.0)
                out_buf = io.BytesIO()
                pil_img.save(out_buf, format="JPEG", quality=85)
                out_buf.seek(0)
                return RLImage(out_buf, width=w * ratio, height=h * ratio)
            except Exception:
                pass

        # 3. Buscar en archivo local foto_path
        foto_path = ev.get("foto_path", "")
        if foto_path and os.path.exists(foto_path):
            try:
                pil_img = PILImage.open(foto_path).convert("RGB")
                w, h = pil_img.size
                ratio = min(max_w / w, max_h / h, 1.0)
                return RLImage(foto_path, width=w * ratio, height=h * ratio)
            except Exception:
                pass

        # 4. Probar en carpeta uploads
        if foto_path:
            cand = UPLOADS_DIR / os.path.basename(foto_path)
            if cand.exists():
                try:
                    pil_img = PILImage.open(str(cand)).convert("RGB")
                    w, h = pil_img.size
                    ratio = min(max_w / w, max_h / h, 1.0)
                    return RLImage(str(cand), width=w * ratio, height=h * ratio)
                except Exception:
                    pass

        # 5. Probar foto_url
        foto_url = ev.get("foto_url", "")
        if foto_url and "/static/uploads/" in foto_url:
            nom = foto_url.split("/static/uploads/")[-1]
            cand = UPLOADS_DIR / nom
            if cand.exists():
                try:
                    pil_img = PILImage.open(str(cand)).convert("RGB")
                    w, h = pil_img.size
                    ratio = min(max_w / w, max_h / h, 1.0)
                    return RLImage(str(cand), width=w * ratio, height=h * ratio)
                except Exception:
                    pass

        return None

    def generar_pdf_inspeccion(self, inspeccion: Dict[str, Any]) -> bytes:
        """
        Genera el informe pericial en PDF estrictamente adaptado a los datos y fotografías
        reales de la inspección auditada.
        """
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            leftMargin=36,
            rightMargin=36,
            topMargin=36,
            bottomMargin=48
        )
        story = []

        # Extraer metadatos de la inspección
        responsable = inspeccion.get("responsable", "Auditor Especialista SST")
        if not responsable or responsable.strip() in ["Auditor SST", "YO", ""]:
            responsable = "Auditor Especialista SST"

        sede = inspeccion.get("sede", "Campus Principal")
        institucion = f"Politécnico Grancolombiano · {sede}" if "Politécnico" not in sede else sede
        edificio = inspeccion.get("edificio", "Edificio Principal")
        area = inspeccion.get("area", "Área de Inspección Técnica")
        fecha = inspeccion.get("fecha", datetime.now().strftime("%d/%m/%Y"))
        numero = inspeccion.get("numero", "#0001")

        evidencias = inspeccion.get("evidencias", [])
        total_evidencias = len(evidencias)

        # Aplanar todos los hallazgos validados o registrados
        elementos_evaluados = []
        for ev in evidencias:
            hallazgos_ev = ev.get("hallazgos", [])
            if hallazgos_ev:
                for h in hallazgos_ev:
                    elementos_evaluados.append({"evidencia": ev, "hallazgo": h})
            else:
                # Si la evidencia no desglosó hallazgos, crear uno con sus campos
                elementos_evaluados.append({
                    "evidencia": ev,
                    "hallazgo": {
                        "elemento": ev.get("tipo", "Elemento de Señalización SST"),
                        "categoria": "General",
                        "condicion": "Evaluado en recorrido",
                        "ubicacion": ev.get("ubicacion", area),
                        "observacion": ev.get("observacion", ""),
                        "norma": "NTC 1461 / Res. 2400 de 1979",
                        "recomendacion": "Garantizar cumplimiento normativo y visibilidad permanente."
                    }
                })

        total_hallazgos = len(elementos_evaluados)

        # -----------------------------------------------------------------
        # ENCABEZADO Y TÍTULO INSTITUCIONAL
        # -----------------------------------------------------------------
        story.append(Paragraph(
            "Informe Pericial de Inspección Técnica de Seguridad y Salud en el Trabajo (SST): "
            "Evaluación de Condiciones Locativas, Sistemas de Emergencia, Señalización y Demarcación",
            self.estilo_titulo
        ))

        meta_parrafo = (
            f"<b>Preparado por:</b> {responsable} &nbsp;|&nbsp; <b>Rol:</b> Especialista / Auditor SST<br/>"
            f"<b>Institución:</b> {institucion} &nbsp;|&nbsp; <b>Edificio:</b> {edificio} &nbsp;|&nbsp; <b>Área Auditada:</b> {area}<br/>"
            f"<b>Fecha de Recorrido:</b> {fecha} &nbsp;|&nbsp; <b>Expediente Correlativo:</b> {numero} &nbsp;|&nbsp; "
            f"<b>Evidencias Documentadas:</b> {total_evidencias} fotográfica(s)"
        )
        story.append(Paragraph(meta_parrafo, self.estilo_subtitulo_meta))
        story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1e3a8a"), spaceAfter=6))

        # -----------------------------------------------------------------
        # 1. INTRODUCCIÓN Y ALCANCE DEL RECORRIDO TÉCNICO
        # -----------------------------------------------------------------
        story.append(Paragraph("1. Introducción y Alcance del Recorrido Técnico", self.estilo_seccion_num))
        intro_texto = (
            f"El presente informe documenta los resultados periciales de la auditoría física y verificación técnica de Seguridad y Salud "
            f"en el Trabajo (SST) llevada a cabo en <b>{institucion}</b>, específicamente en el <b>{edificio}</b> ({area}). "
            f"La auditoría tuvo como propósito evaluar las condiciones reales de seguridad locativa, señalización reglamentaria, equipos contra incendio "
            f"y medios de evacuación, contrastándolos rigurosamente contra el marco legal colombiano (Decreto 1072 de 2015, Resolución 2400 de 1979, "
            f"NSR-10 Títulos J y K, y NTC 1461). Durante el recorrido técnico se procesaron y documentaron pericialmente <b>{total_evidencias} evidencia(s) "
            f"fotográfica(s)</b> correspondientes a <b>{total_hallazgos} elemento(s) normativo(s)</b>, garantizando la trazabilidad técnica exigida por el SG-SST."
        )
        story.append(Paragraph(intro_texto, self.estilo_parrafo))
        story.append(Spacer(1, 2))

        # -----------------------------------------------------------------
        # 2. MARCO LEGAL Y NORMAS TÉCNICAS OBLIGATORIAS DE REFERENCIA
        # -----------------------------------------------------------------
        story.append(Paragraph("2. Marco Legal y Normas Técnicas Obligatorias de Referencia", self.estilo_seccion_num))
        story.append(Paragraph(
            "La evaluación técnica se fundamenta en los siguientes estándares de obligatorio cumplimiento en Colombia:",
            self.estilo_parrafo
        ))

        normas = [
            ("NTC 1461 / NTC-ISO 3864-1 y 3:",
             "Principios cromáticos de seguridad (rojo, amarillo, azul, verde), colores de contraste, formas geométricas y porcentajes mínimos de superficie coloreada."),
            ("NTC-ISO 7010:2021:",
             "Catálogo internacional y nacional de símbolos gráficos registrados para señales de seguridad (prohibición, advertencia, obligación y salvamento)."),
            ("Resolución 2400 de 1979 (Estatuto de Seguridad Industrial):",
             "<b>Art. 17:</b> Pisos libres de humedad y sustancias resbalosas. <b>Art. 164:</b> Dotación obligatoria de camillas y botiquines. "
             "<b>Art. 182:</b> Restricción de alimentos en áreas técnicas y laboratorios. <b>Art. 281:</b> Pasillos con demarcación continua amarilla de 10 cm y ancho libre >= 40 cm."),
            ("NSR-10 (Reglamento Colombiano de Construcción Sismo Resistente):",
             "• <b>Título J (Protección contra Incendios):</b> Dotación de extintores portátiles según carga combustible (NTC 2885 / NFPA 10), altura máx. 1.50m y red de gabinetes (NTC 1669 / NFPA 14).<br/>"
             "• <b>Título K (Medios de Evacuación):</b> Dimensionamiento de pasillos de escape, puertas de emergencia y señalética direccional continua hacia salidas."),
            ("ISO 16069 (Sistemas SWGS):",
             "Señalización fotoluminiscente en tres planos: alto (>2.00m), intermedio (1.60m-1.80m) y zócalo (10-40cm del piso) para escape bajo humo."),
            ("Decreto 1072 de 2015 (Art. 2.2.4.6.25):",
             "Implementación y dotación técnica del Plan de Prevención, Preparación y Respuesta ante Emergencias de la organización."),
            ("Ley 1335 de 2009 (Ley Antitabaco - Arts. 19-20):",
             "Garantía de espacios 100% libres de humo de tabaco y colocación obligatoria de avisos restrictivos en todos los accesos.")
        ]

        for tit, desc in normas:
            story.append(Paragraph(f"• <b>{tit}</b> {desc}", self.estilo_norma_item))

        story.append(Spacer(1, 3))

        # -----------------------------------------------------------------
        # 3. GUÍA METROLÓGICA Y CROMÁTICA DE SEÑALIZACIÓN Y DEMARCACIÓN
        # -----------------------------------------------------------------
        story.append(Paragraph("3. Guía Metrológica y Cromática de Señalización y Demarcación", self.estilo_seccion_num))
        story.append(Paragraph(
            "Conforme a la NTC 1461 y NTC-ISO 3864, toda señalización debe cumplir parámetros cromáticos y dimensionamiento metrológico exacto:",
            self.estilo_parrafo
        ))

        # Tabla cromática
        tabla_cromat_datos = [
            [
                Paragraph("<b>Color Seguridad</b>", self.estilo_celda_th),
                Paragraph("<b>Contraste</b>", self.estilo_celda_th),
                Paragraph("<b>Forma Geométrica</b>", self.estilo_celda_th),
                Paragraph("<b>% Mínimo</b>", self.estilo_celda_th),
                Paragraph("<b>Significado Técnico y Aplicación</b>", self.estilo_celda_th)
            ],
            [
                Paragraph("<font color='#dc2626'><b>ROJO</b></font>", self.estilo_celda_bold),
                Paragraph("Blanco", self.estilo_celda),
                Paragraph("Círculo con barra diagonal descendente a 45°", self.estilo_celda),
                Paragraph(">= 35%", self.estilo_celda_center),
                Paragraph("<b>Prohibición:</b> Prohibido fumar (Ley 1335/09), no celular, no alimentos, parada de emergencia.", self.estilo_celda)
            ],
            [
                Paragraph("<font color='#dc2626'><b>ROJO</b></font>", self.estilo_celda_bold),
                Paragraph("Blanco", self.estilo_celda),
                Paragraph("Cuadrado o Rectángulo", self.estilo_celda),
                Paragraph(">= 50%", self.estilo_celda_center),
                Paragraph("<b>Equipos Contra Incendio:</b> Extintores portátiles, gabinetes, pulsadores manuales de alarma.", self.estilo_celda)
            ],
            [
                Paragraph("<font color='#d97706'><b>AMARILLO</b></font>", self.estilo_celda_bold),
                Paragraph("Negro", self.estilo_celda),
                Paragraph("Triángulo equilátero / Franjas a 45°", self.estilo_celda),
                Paragraph(">= 50%", self.estilo_celda_center),
                Paragraph("<b>Advertencia / Riesgo:</b> Piso resbaloso (Res. 2400/79 Art. 17), riesgo eléctrico, desniveles.", self.estilo_celda)
            ],
            [
                Paragraph("<font color='#2563eb'><b>AZUL</b></font>", self.estilo_celda_bold),
                Paragraph("Blanco", self.estilo_celda),
                Paragraph("Círculo pleno", self.estilo_celda),
                Paragraph(">= 50%", self.estilo_celda_center),
                Paragraph("<b>Obligación:</b> Uso imperativo de EPP (casco, gafas, calzado de seguridad, protector auditivo).", self.estilo_celda)
            ],
            [
                Paragraph("<font color='#059669'><b>VERDE</b></font>", self.estilo_celda_bold),
                Paragraph("Blanco", self.estilo_celda),
                Paragraph("Cuadrado o Rectángulo", self.estilo_celda),
                Paragraph(">= 50%", self.estilo_celda_center),
                Paragraph("<b>Salvamento / Evacuación:</b> Salidas de emergencia, camillas rígidas, botiquín, punto de encuentro.", self.estilo_celda)
            ]
        ]
        t_cromat = Table(tabla_cromat_datos, colWidths=[70, 55, 120, 50, 245])
        t_cromat.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0f2444')),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#94a3b8')),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor('#ffffff'), colors.HexColor('#f8fafc')]),
            ('TOPPADDING', (0,0), (-1,-1), 2.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
            ('LEFTPADDING', (0,0), (-1,-1), 3),
            ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ]))
        story.append(t_cromat)
        story.append(Spacer(1, 4))

        # Tabla de dimensionamiento metrológico (S >= d^2 / 2000)
        story.append(Paragraph("<b>Dimensionamiento y Distancias de Observación (Fórmula Legal S &ge; d² / 2000):</b>", self.estilo_parrafo))
        tabla_dim_datos = [
            [
                Paragraph("<b>Distancia Máx. (d)</b>", self.estilo_celda_th),
                Paragraph("<b>Superficie Mín. (S)</b>", self.estilo_celda_th),
                Paragraph("<b>Dimensión Cuadrada</b>", self.estilo_celda_th),
                Paragraph("<b>Dimensión Rectangular</b>", self.estilo_celda_th),
                Paragraph("<b>Aplicación en Campo</b>", self.estilo_celda_th)
            ],
            [
                Paragraph("Hasta 5 metros", self.estilo_celda_bold),
                Paragraph("&ge; 0.0125 m²", self.estilo_celda_center),
                Paragraph("10 x 10 cm", self.estilo_celda_center),
                Paragraph("15 x 10 cm", self.estilo_celda_center),
                Paragraph("Tableros eléctricos individuales, puestos puntuales.", self.estilo_celda)
            ],
            [
                Paragraph("Hasta 10 metros", self.estilo_celda_bold),
                Paragraph("&ge; 0.0500 m²", self.estilo_celda_center),
                Paragraph("20 x 20 cm", self.estilo_celda_center),
                Paragraph("30 x 15 cm / 30 x 20 cm", self.estilo_celda_center),
                Paragraph("Pasillos estándar de laboratorios, extintores de pared.", self.estilo_celda)
            ],
            [
                Paragraph("Hasta 15 metros", self.estilo_celda_bold),
                Paragraph("&ge; 0.1125 m²", self.estilo_celda_center),
                Paragraph("30 x 30 cm", self.estilo_celda_center),
                Paragraph("40 x 20 cm / 45 x 20 cm", self.estilo_celda_center),
                Paragraph("Corredores principales, accesos a bloques, gabinetes hídricos.", self.estilo_celda)
            ],
            [
                Paragraph("Hasta 20 metros", self.estilo_celda_bold),
                Paragraph("&ge; 0.2000 m²", self.estilo_celda_center),
                Paragraph("42 x 42 cm (DIN A3)", self.estilo_celda_center),
                Paragraph("60 x 30 cm", self.estilo_celda_center),
                Paragraph("Auditorios, cafeterías, zonas comunes amplias.", self.estilo_celda)
            ],
            [
                Paragraph("Hasta 30 metros", self.estilo_celda_bold),
                Paragraph("&ge; 0.4500 m²", self.estilo_celda_center),
                Paragraph("60 x 60 cm (DIN A2)", self.estilo_celda_center),
                Paragraph("80 x 40 cm", self.estilo_celda_center),
                Paragraph("Patios exteriores, zonas de parqueo y punto de encuentro.", self.estilo_celda)
            ]
        ]
        t_dim = Table(tabla_dim_datos, colWidths=[85, 80, 85, 100, 190])
        t_dim.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1e3a8a')),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#94a3b8')),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor('#ffffff'), colors.HexColor('#f8fafc')]),
            ('TOPPADDING', (0,0), (-1,-1), 2.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
            ('LEFTPADDING', (0,0), (-1,-1), 3),
            ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ]))
        story.append(t_dim)
        story.append(Spacer(1, 4))

        # Cotas reglamentarias
        texto_cotas = (
            "<b>Cotas Reglamentarias de Montaje (Res. 2400/79 e ISO 16069):</b> "
            "• Altura estándar pared: 1.60m a 1.80m al centro de la señal. "
            "• Suspendidas/bandera: &ge; 2.00m libre. "
            "• Nivel de zócalo (ISO 16069): 10cm a 40cm del piso. "
            "• Extintores (NTC 2885): Manija a máx. 1.50m (peso &le; 18kg) o 1.10m (> 18kg), cuadro de despeje 1.0x1.0m con franjas amarillo/negro de 10cm. "
            "• Calidad fotoluminiscente: Clase A (&ge; 210 mcd/m² a 10 min) en edificios de concurrencia pública."
        )
        story.append(Paragraph(texto_cotas, self.estilo_norma_item))
        story.append(Spacer(1, 4))

        # -----------------------------------------------------------------
        # 4. MATRIZ DE HALLAZGOS, EVALUACIÓN DE CONFORMIDAD Y PARÁMETROS TÉCNICOS
        # (DINÁMICA: Basada exclusivamente en los hallazgos de ESTA inspección)
        # -----------------------------------------------------------------
        story.append(Paragraph("4. Matriz de Hallazgos, Evaluación de Conformidad y Parámetros Técnicos", self.estilo_seccion_num))
        story.append(Paragraph(
            "A continuación, se detalla la matriz pericial con la evaluación técnica de conformidad de los elementos inspeccionados "
            "en el presente recorrido, especificando la norma aplicable, parámetros cromáticos, tamaño reglamentario y nivel de riesgo:",
            self.estilo_parrafo
        ))

        tabla_matriz_filas = [
            [
                Paragraph("<b>Ítem / Elemento</b>", self.estilo_celda_th),
                Paragraph("<b>Ubicación y Condición Observada</b>", self.estilo_celda_th),
                Paragraph("<b>Norma Técnica</b>", self.estilo_celda_th),
                Paragraph("<b>Especificación (Color, S&ge;d²/2000, Altura)</b>", self.estilo_celda_th),
                Paragraph("<b>Riesgo</b>", self.estilo_celda_th),
                Paragraph("<b>Recomendación Técnica Oficial</b>", self.estilo_celda_th),
            ]
        ]

        conformes_count = 0
        subestandar_count = 0

        for idx, item in enumerate(elementos_evaluados, 1):
            h = item["hallazgo"]
            ev = item["evidencia"]

            elem_nombre = h.get("elemento") or ev.get("tipo") or f"Elemento #{idx}"
            ubi = h.get("ubicacion") or ev.get("ubicacion") or area
            cond = h.get("condicion") or "Requiere verificación"
            obs = h.get("observacion") or ev.get("observacion") or "Condición observada durante el recorrido técnico."
            norma_val = h.get("norma") or "NTC 1461 / NSR-10 / Res. 2400 de 1979"
            recom_val = h.get("recomendacion") or "Garantizar mantenimiento y verificación periódica conforme a norma técnica."
            
            # Especificaciones
            col_seg = h.get("color_seguridad") or "Normalizado"
            tam_rec = h.get("tamano_recomendado") or "20x20 cm (d<=10m)"
            alt_mon = h.get("altura_montaje") or "1.60m - 1.80m"
            especif_txt = f"{col_seg} | Tamaño: {tam_rec} | Montaje: {alt_mon}"

            # Evaluar nivel de riesgo y conformidad
            cond_lower = cond.lower()
            if "conforme" in cond_lower and "no" not in cond_lower and "requiere" not in cond_lower:
                conformes_count += 1
                nivel_riesgo = "CONTROLADO"
                color_riesgo = "#059669"
            else:
                subestandar_count += 1
                cat_lower = (h.get("categoria") or "").lower()
                elem_lower = elem_nombre.lower()
                if any(k in elem_lower or k in cat_lower for k in ["incendio", "extintor", "gabinete", "evacuacion", "salida", "camilla"]):
                    nivel_riesgo = "ALTO"
                    color_riesgo = "#dc2626"
                elif any(k in elem_lower or k in cat_lower for k in ["resbaloso", "caida", "humed", "advertencia"]):
                    nivel_riesgo = "MEDIO"
                    color_riesgo = "#d97706"
                else:
                    nivel_riesgo = "MEDIO"
                    color_riesgo = "#2563eb"

            p_item = Paragraph(f"<b>{idx}. {elem_nombre}</b>", self.estilo_celda_bold)
            p_cond = Paragraph(f"<b>{ubi}</b><br/>Condición: <i>{cond}</i><br/>{obs}", self.estilo_celda)
            p_norm = Paragraph(norma_val, self.estilo_celda)
            p_espec = Paragraph(especif_txt, self.estilo_celda)
            p_riesgo = Paragraph(f"<font color='{color_riesgo}'><b>{nivel_riesgo}</b></font>", self.estilo_celda_bold)
            p_recom = Paragraph(recom_val, self.estilo_celda)

            tabla_matriz_filas.append([p_item, p_cond, p_norm, p_espec, p_riesgo, p_recom])

        if not elementos_evaluados:
            tabla_matriz_filas.append([
                Paragraph("<b>Sin registros</b>", self.estilo_celda),
                Paragraph("No se registraron hallazgos individuales en este recorrido.", self.estilo_celda),
                Paragraph("NTC 1461", self.estilo_celda),
                Paragraph("---", self.estilo_celda),
                Paragraph("BAJO", self.estilo_celda),
                Paragraph("Mantener rutinas periódicas de inspección preventiva.", self.estilo_celda)
            ])

        t_matriz = Table(tabla_matriz_filas, colWidths=[75, 110, 95, 110, 50, 100])
        t_matriz.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0f2444')),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#94a3b8')),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor('#ffffff'), colors.HexColor('#f8fafc')]),
            ('TOPPADDING', (0,0), (-1,-1), 2.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
            ('LEFTPADDING', (0,0), (-1,-1), 3),
            ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ]))
        story.append(t_matriz)
        story.append(Spacer(1, 5))

        # -----------------------------------------------------------------
        # 5. EVIDENCIAS Y REGISTRO FOTOGRÁFICO PERICIAL (ESTÁNDAR APA 7ma EDICIÓN)
        # (DINÁMICA: Muestra exactamente las fotografías de ESTA inspección)
        # -----------------------------------------------------------------
        story.append(Paragraph("5. Evidencias y Registro Fotográfico Pericial (Estándar APA 7ma Edición)", self.estilo_seccion_num))
        story.append(Paragraph(
            "<i>(Nota metodológica: Las evidencias fotográficas capturadas en el recorrido se estructuran bajo el estándar editorial "
            "APA 7ma edición. Para cada figura se presenta número correlativo, título en cursiva, imagen fotográfica y nota pericial "
            "con especificación de normas, dimensiones y cotas de altura).</i>",
            self.estilo_parrafo
        ))

        if not evidencias:
            story.append(Paragraph("<i>No se adjuntaron fotografías en este expediente de auditoría.</i>", self.estilo_celda))
        else:
            for num_fig, ev in enumerate(evidencias, 1):
                # Extraer datos de la evidencia
                h_primero = ev.get("hallazgos", [{}])[0] if ev.get("hallazgos") else {}
                elem_nom = h_primero.get("elemento") or ev.get("tipo") or f"Evidencia Técnica #{num_fig}"
                ubi_ev = h_primero.get("ubicacion") or ev.get("ubicacion") or area
                cond_ev = h_primero.get("condicion") or "Registrado"
                norma_ev = h_primero.get("norma") or "NTC 1461 / Res. 2400 de 1979 / NSR-10"
                obs_ev = ev.get("observacion") or h_primero.get("observacion") or "Evidencia fotográfica documentada en terreno."
                recom_ev = h_primero.get("recomendacion") or "Verificar adecuación técnica según parámetros normativos."
                alt_ev = h_primero.get("altura_montaje") or "1.60m - 1.80m"
                tam_ev = h_primero.get("tamano_recomendado") or "20x20 cm (S >= d²/2000)"

                bloque_figura = []
                bloque_figura.append(Paragraph(f"<b>Figura {num_fig}</b>", self.estilo_apa_figura))
                bloque_figura.append(Paragraph(f"<i>Registro pericial de {elem_nom} en {ubi_ev}</i>", self.estilo_apa_titulo))

                rl_img = self._cargar_imagen_flowable(ev)
                if rl_img:
                    t_img = Table([[rl_img]], colWidths=[540], style=[
                        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                        ('LEFTPADDING', (0,0), (-1,-1), 0),
                        ('RIGHTPADDING', (0,0), (-1,-1), 0),
                        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
                    ])
                    bloque_figura.append(t_img)
                else:
                    p_box = Table(
                        [[Paragraph(
                            f"<b>[Registro Fotográfico - Figura {num_fig}: {elem_nom}]</b><br/>"
                            f"<font size=7 color='#475569'>Evidencia documentada en: {ubi_ev}. Condición: {cond_ev}.</font><br/>"
                            f"<font size=6.5 color='#64748b'>Verificación de conformidad con NTC 1461 y NSR-10.</font>",
                            self.estilo_celda_center
                        )]],
                        colWidths=[540],
                        style=[
                            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f8fafc')),
                            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#cbd5e1')),
                            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                            ('TOPPADDING', (0,0), (-1,-1), 10),
                            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
                        ]
                    )
                    bloque_figura.append(p_box)

                nota_completa = (
                    f"<i>Nota.</i> Condición evaluada: <b>{cond_ev}</b>. {obs_ev} "
                    f"Marco legal: {norma_ev}. Parámetros técnicos exigidos: dimensión {tam_ev}, altura de montaje reglamentaria a {alt_ev}. "
                    f"Dictamen técnico: {recom_ev}"
                )
                bloque_figura.append(Paragraph(nota_completa, self.estilo_apa_nota))
                story.append(KeepTogether(bloque_figura))

        story.append(Spacer(1, 4))

        # -----------------------------------------------------------------
        # 6. PLAN DE ACCIÓN PRIORIZADO Y CRONOGRAMA DE CUMPLIMIENTO
        # (DINÁMICO: Formulado a partir de las necesidades reales de esta inspección)
        # -----------------------------------------------------------------
        story.append(Paragraph("6. Plan de Acción Priorizado y Cronograma de Cumplimiento", self.estilo_seccion_num))
        story.append(Paragraph(
            "Con base en la matriz de hallazgos evaluada, se establece el siguiente plan de intervención técnica priorizado "
            "para garantizar el cierre de brechas y la conformidad con los estándares del Consejo Colombiano de Seguridad (CCS) y la NTC 1461:",
            self.estilo_parrafo
        ))

        # Generar actividades basadas en los elementos reales
        tiene_incendio = any("incendio" in (i["hallazgo"].get("elemento","") + i["hallazgo"].get("categoria","")).lower() or "extintor" in (i["hallazgo"].get("elemento","")).lower() for i in elementos_evaluados)
        tiene_evacuacion = any("evacuacion" in (i["hallazgo"].get("elemento","") + i["hallazgo"].get("categoria","")).lower() or "salida" in (i["hallazgo"].get("elemento","")).lower() for i in elementos_evaluados)
        tiene_primeros_auxilios = any("camilla" in (i["hallazgo"].get("elemento","")).lower() or "botiquin" in (i["hallazgo"].get("elemento","")).lower() for i in elementos_evaluados)
        tiene_advertencia = any("resbaloso" in (i["hallazgo"].get("elemento","")).lower() or "caida" in (i["hallazgo"].get("elemento","")).lower() or "advertencia" in (i["hallazgo"].get("categoria","")).lower() for i in elementos_evaluados)

        actividades_fase1 = []
        if tiene_incendio:
            actividades_fase1.append("Verificación de presión en manómetros de extintores (zona verde), despeje frontal de 1.0m y tarjetas de inspección mensual.")
        if tiene_primeros_auxilios:
            actividades_fase1.append("Dotación de camilla rígida de rescate con inmovilizadores y revisión de insumos del botiquín.")
        if tiene_advertencia:
            actividades_fase1.append("Despliegue obligatorio de caballetes tipo tijera de 60cm de piso resbaloso durante rutinas de aseo.")
        if not actividades_fase1:
            actividades_fase1.append("Subsanación de condiciones inmediatas identificadas en la matriz y verificación de despeje de pasillos.")

        actividades_fase2 = []
        if tiene_evacuacion:
            actividades_fase2.append("Instalación de señales fotoluminiscentes Clase A (210 mcd/m²) en rutas de evacuación cada 15m a 1.70m y a nivel de zócalo a 30cm.")
        actividades_fase2.append("Demarcación continua de franjas amarillas de 10 cm en pasillos (Res. 2400/79 Art. 281) y fijación de avisos restrictivos Ley 1335.")

        actividades_fase3 = [
            "Auditoría de seguimiento y cierre formal de hallazgos en la plataforma SST-IA.",
            "Actualización de la planimetría de emergencias institucional conforme al Decreto 1072 de 2015."
        ]

        tabla_plan_datos = [
            [
                Paragraph("<b>Fase / Plazo</b>", self.estilo_celda_th),
                Paragraph("<b>Actividad Técnica y Especificación</b>", self.estilo_celda_th),
                Paragraph("<b>Norma Reguladora</b>", self.estilo_celda_th),
                Paragraph("<b>Responsable</b>", self.estilo_celda_th),
                Paragraph("<b>Entregable Verificable</b>", self.estilo_celda_th)
            ],
            [
                Paragraph("<b>Fase 1: Inmediato</b><br/>(&lt; 48 Horas)", self.estilo_celda_bold),
                Paragraph("<br/>• ".join(["• " + a for a in actividades_fase1]), self.estilo_celda),
                Paragraph("NSR-10 Título J<br/>Res. 2400/79", self.estilo_celda),
                Paragraph(f"Auditor SST ({responsable}) / Servicios Generales", self.estilo_celda),
                Paragraph("Acta de inspección y registro fotográfico de verificación inmediata.", self.estilo_celda)
            ],
            [
                Paragraph("<b>Fase 2: Corto Plazo</b><br/>(&lt; 15 Días)", self.estilo_celda_bold),
                Paragraph("<br/>• ".join(["• " + a for a in actividades_fase2]), self.estilo_celda),
                Paragraph("NTC 1461<br/>ISO 16069<br/>Ley 1335/09", self.estilo_celda),
                Paragraph("Mantenimiento Locativo / Planta Física", self.estilo_celda),
                Paragraph("Certificado de instalación fotoluminiscente y demarcación conforme.", self.estilo_celda)
            ],
            [
                Paragraph("<b>Fase 3: Mediano Plazo</b><br/>(&lt; 30 Días)", self.estilo_celda_bold),
                Paragraph("<br/>• ".join(["• " + a for a in actividades_fase3]), self.estilo_celda),
                Paragraph("Decreto 1072/15<br/>NSR-10 Título K", self.estilo_celda),
                Paragraph("Coordinación de SST / Gerencia", self.estilo_celda),
                Paragraph("Informe de cierre de ciclo en el SG-SST institucional.", self.estilo_celda)
            ]
        ]

        t_plan = Table(tabla_plan_datos, colWidths=[80, 165, 95, 100, 100])
        t_plan.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0f2444')),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#94a3b8')),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor('#ffffff'), colors.HexColor('#f8fafc')]),
            ('TOPPADDING', (0,0), (-1,-1), 2.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
            ('LEFTPADDING', (0,0), (-1,-1), 3),
            ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ]))
        story.append(t_plan)
        story.append(Spacer(1, 4))

        # -----------------------------------------------------------------
        # 7. CONCLUSIONES Y DICTAMEN DE AUDITORÍA PERICIAL
        # (DINÁMICO: Estadísticas reales y dictamen fundamentado)
        # -----------------------------------------------------------------
        conclusiones_bloque = []
        conclusiones_bloque.append(Paragraph("7. Conclusiones y Dictamen de Auditoría Pericial", self.estilo_seccion_num))

        pct_conformidad = int((conformes_count / total_hallazgos) * 100) if total_hallazgos > 0 else 100
        
        if subestandar_count > 0:
            dictamen_eval = (
                f"El recorrido técnico constató la existencia de condiciones subestándar que ameritan intervención correctiva prioritaria. "
                f"De un total de <b>{total_hallazgos} parámetro(s) normativo(s)</b> evaluados sobre <b>{total_evidencias} evidencia(s) fotográfica(s)</b>, "
                f"se verificó un índice de conformidad del <b>{pct_conformidad}%</b> ({conformes_count} ítem(s) conformes y {subestandar_count} ítem(s) con requerimiento técnico). "
                f"El incumplimiento de las especificaciones de señalización fotoluminiscente (NTC 1461 / ISO 16069), alturas de fijación y protección contra incendios (NSR-10 Título J) "
                f"eleva la vulnerabilidad de las instalaciones ante una contingencia. Se dictamina ejecutar con carácter estricto el cronograma del numeral 6."
            )
        else:
            dictamen_eval = (
                f"La auditoría técnica constata que los elementos evaluados en el área <b>{area}</b> de <b>{institucion}</b> presentan un índice "
                f"de conformidad del <b>{pct_conformidad}%</b> frente a las exigencias de la NTC 1461, NSR-10 y Resolución 2400 de 1979. "
                f"Se recomienda sostener las inspecciones técnicas periódicas mensuales para garantizar la vigencia operativa continua."
            )

        conclusiones_bloque.append(Paragraph(dictamen_eval, self.estilo_parrafo))
        conclusiones_bloque.append(Spacer(1, 16))

        # Firmas periciales
        conclusiones_bloque.append(Table([
            [
                Paragraph(
                    "____________________________________________________<br/>"
                    f"<b>{responsable}</b><br/>"
                    "Auditor Líder / Especialista en Seguridad y Salud en el Trabajo<br/>"
                    "Licencia en Salud Ocupacional / SST Vigente<br/>"
                    f"{institucion}",
                    self.estilo_celda
                ),
                Paragraph(
                    "____________________________________________________<br/>"
                    "<b>Rectoría / Coordinación de SST y Planta Física</b><br/>"
                    "Representante Legal / Gerencia del SG-SST<br/>"
                    f"{institucion}",
                    self.estilo_celda
                )
            ]
        ], colWidths=[270, 270], style=[
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'TOP')
        ]))

        story.append(KeepTogether(conclusiones_bloque))

        doc.build(story, canvasmaker=NumberedCanvas)
        return buffer.getvalue()


pdf_service = PDFService()
