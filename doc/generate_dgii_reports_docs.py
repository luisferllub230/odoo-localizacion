#!/usr/bin/env python3
"""Generate technical and user manual Word documents for dgii_reports module."""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy


# ─── Color constants ──────────────────────────────────────────────────────────
NAVY      = RGBColor(0x1F, 0x38, 0x64)
BLUE1     = RGBColor(0x36, 0x5F, 0x91)
BLUE2     = RGBColor(0x4F, 0x81, 0xBD)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
GRAY_LIGHT = RGBColor(0xD9, 0xE2, 0xF3)
IMG_BG    = RGBColor(0xED, 0xED, 0xED)


# ─── Helpers ──────────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color: str):
    """Set table cell background color."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    existing = tcPr.find(qn('w:shd'))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(shd)


def add_table(doc, headers, rows, col_widths=None):
    """Add a formatted table with navy header row."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    # Header row
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        set_cell_bg(hdr_cells[i], '1F3864')
        for run in hdr_cells[i].paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = WHITE
            run.font.size = Pt(9)
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        hdr_cells[i].paragraphs[0].runs[0].font.color.rgb = WHITE
    # Data rows
    for ri, row_data in enumerate(rows):
        cells = table.rows[ri + 1].cells
        for ci, val in enumerate(row_data):
            cells[ci].text = str(val)
            for run in cells[ci].paragraphs[0].runs:
                run.font.size = Pt(9)
            if ri % 2 == 0:
                set_cell_bg(cells[ci], 'EBF3FB')
    # Column widths
    if col_widths:
        for row in table.rows:
            for ci, w in enumerate(col_widths):
                row.cells[ci].width = Inches(w)
    doc.add_paragraph()
    return table


def add_image_placeholder(doc, caption: str):
    """Add a grey box placeholder for an image."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'[ IMAGEN: {caption} ]')
    run.bold = True
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    run.font.size = Pt(10)
    # Gray shading on paragraph
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'EDEDEB')
    pPr.append(shd)
    # Space above/below
    pf = p.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    caption_p = doc.add_paragraph(caption)
    caption_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in caption_p.runs:
        run.italic = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    return p


def add_cover(doc, title_line1, title_line2, subtitle, doc_number, doc_type_label):
    """Add cover page paragraphs."""
    for _ in range(6):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title_line1)
    run.bold = True
    run.font.size = Pt(26)
    run.font.color.rgb = NAVY

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(title_line2)
    run2.bold = True
    run2.font.size = Pt(16)
    run2.font.color.rgb = NAVY

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run3 = p3.add_run(subtitle)
    run3.bold = True
    run3.font.size = Pt(12)
    run3.font.color.rgb = BLUE1

    doc.add_paragraph()

    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run4 = p4.add_run(doc_type_label)
    run4.bold = True
    run4.font.size = Pt(11)
    run4.font.color.rgb = NAVY

    p5 = doc.add_paragraph()
    p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run5 = p5.add_run(doc_number)
    run5.font.size = Pt(10)
    run5.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    doc.add_page_break()


def heading1(doc, text):
    return doc.add_heading(text, level=1)


def heading2(doc, text):
    return doc.add_heading(text, level=2)


def heading3(doc, text):
    return doc.add_heading(text, level=3)


def bullet(doc, text):
    return doc.add_paragraph(text, style='List Bullet')


def numbered(doc, text):
    return doc.add_paragraph(text, style='List Number')


def normal(doc, text):
    return doc.add_paragraph(text)


def code_block(doc, text):
    p = doc.add_paragraph(text)
    p.style = doc.styles['No Spacing']
    for run in p.runs:
        run.font.name = 'Courier New'
        run.font.size = Pt(9)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'F2F2F2')
    pPr.append(shd)
    return p


# ═══════════════════════════════════════════════════════════════════════════════
#  TECHNICAL DOCUMENT
# ═══════════════════════════════════════════════════════════════════════════════

def build_technical_doc():
    doc = Document('doc/doc_tecnico_l10n_do_accounting.docx')
    # Start fresh keeping styles — preserve sectPr (page layout)
    body = doc.element.body
    sectPr = body.find(qn('w:sectPr'))
    for element in list(body):
        body.remove(element)
    if sectPr is not None:
        body.append(sectPr)

    # ── Cover ─────────────────────────────────────────────────────────────────
    add_cover(
        doc,
        'Documentación Técnica',
        'Módulo dgii_reports\nDeclaraciones DGII — Odoo 19.0',
        'ESPECIFICACIÓN TÉCNICA DE SOFTWARE',
        'Documento N.° LFL-TEC-002    |    Confidencial — Uso Interno',
        'ESPECIFICACIÓN TÉCNICA DE SOFTWARE',
    )

    # ── TOC ───────────────────────────────────────────────────────────────────
    heading1(doc, 'Tabla de Contenido')
    toc_entries = [
        '1.  Información General del Módulo',
        '2.  Estructura de Directorios',
        '3.  Modelos Principales',
        '    3.1  dgii.reports',
        '    3.2  dgii.reports.purchase.line',
        '    3.3  dgii.reports.sale.line',
        '    3.4  dgii.reports.cancel.line',
        '    3.5  dgii.reports.exterior.line',
        '    3.6  dgii.reports.it1.line',
        '    3.7  invoice.service.type.detail',
        '4.  Extensiones a Modelos Base',
        '    4.1  account.move',
        '    4.2  account.tax',
        '    4.3  account.journal',
        '    4.4  account.account',
        '    4.5  res.company',
        '    4.6  res.partner',
        '5.  Wizards (Modelos Transitorios)',
        '6.  Controladores HTTP',
        '7.  Seguridad y Control de Acceso',
        '8.  Datos Maestros',
        '9.  Hooks de Instalación',
        '    9.1  set_l10n_do_tax_types(cr)',
        '    9.2  set_journal_payment_forms(cr)',
        '    9.3  Datos de Demostración',
        '10. Reglas de Negocio Críticas',
        '11. Formatos de Archivo TXT (DGII)',
        '12. Pruebas Unitarias',
        '13. Configuración del Entorno de Desarrollo',
        '14. Variables de Contexto y Depuración',
    ]
    for e in toc_entries:
        normal(doc, e)
    doc.add_page_break()

    # ── Section 1 ─────────────────────────────────────────────────────────────
    heading1(doc, '1. Información General del Módulo')
    normal(doc, (
        'El presente documento constituye la especificación técnica detallada del módulo dgii_reports, '
        'desarrollado para la localización dominicana de Odoo 19.0. El módulo automatiza la generación '
        'de las declaraciones fiscales obligatorias ante la Dirección General de Impuestos Internos (DGII) '
        'de la República Dominicana, incluyendo los reportes 606, 607, 608, 609, Anexo A e IT-1.'
    ))
    add_table(doc,
        ['Atributo', 'Valor'],
        [
            ['Nombre técnico', 'dgii_reports'],
            ['Nombre visible', 'Declaraciones DGII'],
            ['Versión del módulo', '19.0.1.0.0'],
            ['Autor', 'lfernandez'],
            ['Licencia', 'LGPL-3'],
            ['Versión Odoo', '19.0'],
        ],
        col_widths=[2.0, 4.0]
    )

    heading2(doc, '1.1 Dependencias del Módulo')
    normal(doc, 'El módulo declara las siguientes dependencias directas en su archivo __manifest__.py:')
    add_table(doc,
        ['Módulo dependencia', 'Descripción'],
        [
            ['l10n_do', 'Localización base de República Dominicana. Tipos de documento NCF.'],
            ['l10n_do_accounting', 'Módulo de comprobantes fiscales. NCF Batches, facturas fiscales.'],
            ['accountant', 'Módulo de contabilidad avanzada de Odoo.'],
            ['account', 'Módulo base de contabilidad de Odoo.'],
            ['web', 'Framework web de Odoo.'],
        ],
        col_widths=[2.2, 3.8]
    )
    normal(doc, 'Dependencia externa Python:')
    add_table(doc,
        ['Paquete', 'Uso'],
        [
            ['pycountry', 'Conversión de nombres de países a códigos numéricos ISO 3166 para el reporte 609.'],
        ],
        col_widths=[1.8, 4.2]
    )

    heading2(doc, '1.2 Resumen Funcional')
    normal(doc, 'El módulo provee las siguientes capacidades funcionales principales:')
    for item in [
        'Generación del reporte 606 (Registro de Compras y Gastos) en formato TXT DGII.',
        'Generación del reporte 607 (Registro de Ventas) en formato TXT DGII.',
        'Generación del reporte 608 (Comprobantes Anulados) en formato TXT DGII.',
        'Generación del reporte 609 (Pagos al Exterior) en formato TXT DGII.',
        'Cálculo automático del Anexo A (IT-1) con 56 casillas agrupadas en 6 secciones.',
        'Cálculo automático del formulario IT-1 (Declaración ITBIS) con 68 casillas.',
        'Resumen automático de comprobantes consumidor (B02) para la declaración.',
        'Consolidación multi-sucursal: un reporte consolida datos de compañía padre e hijos.',
        'Marcado de facturas reportadas (fiscal_status = done) al marcar el período como enviado.',
        'Clasificación de impuestos por tipo (ITBIS, RITBIS, ISR, ISC, propina legal).',
        'Soporte para proveedores del exterior (B17) con tipo de servicio y país ISO.',
        'Widget JavaScript para visualización del resumen de declaraciones en el formulario.',
    ]:
        bullet(doc, item)

    doc.add_page_break()

    # ── Section 2 ─────────────────────────────────────────────────────────────
    heading1(doc, '2. Estructura de Directorios')
    normal(doc, 'Árbol de directorios del módulo con descripción funcional de cada componente:')
    code_block(doc, (
        'dgii_reports/\n'
        '├── __init__.py                   Inicialización del paquete Python\n'
        '├── __manifest__.py               Metadatos, dependencias y assets del módulo\n'
        '├── hooks.py                      post_init_hook: tipifica impuestos y diarios\n'
        '├── requirements.txt              Dependencias Python: pycountry\n'
        '├── controllers/\n'
        '│   ├── __init__.py\n'
        '│   └── main.py                   Endpoint HTTP /dgii_reports/<ncf_rnc>\n'
        '├── data/\n'
        '│   └── invoice_service_type_detail_data.xml   Tipos de servicio B17\n'
        '├── demo/\n'
        '│   └── demo_dgii.xml             Invoca dgii.reports._load_demo_data()\n'
        '├── i18n/                         Archivos de traducción .po\n'
        '├── migrations/                   Scripts de migración entre versiones\n'
        '├── models/\n'
        '│   ├── __init__.py\n'
        '│   ├── account_account.py        Extensión: casillas Anexo A e IT-1\n'
        '│   ├── account_journal.py        Extensión: forma de pago del diario\n'
        '│   ├── account_move.py           Extensión: campos DGII en facturas\n'
        '│   ├── account_tax.py            Extensión: tipo de impuesto DGII\n'
        '│   ├── dgii_reports.py           Modelo principal + líneas + wizard\n'
        '│   ├── invoice_service_type_detail.py  Catálogo B17\n'
        '│   └── res_company.py            Extensión: consolidación sucursales\n'
        '│   └── res_partner.py            Extensión: campo relacionado B17\n'
        '├── security/\n'
        '│   └── ir.model.access.csv       Permisos de acceso ORM\n'
        '├── static/\n'
        '│   └── src/\n'
        '│       ├── js/widget.js          Widget dashboard declaraciones\n'
        '│       └── scss/dgii_reports.scss   Estilos del widget\n'
        '├── tests/                        Suite de pruebas unitarias\n'
        '├── views/                        Vistas XML de todos los modelos\n'
        '└── wizard/                       Wizard de regeneración\n'
    ))

    doc.add_page_break()

    # ── Section 3 ─────────────────────────────────────────────────────────────
    heading1(doc, '3. Modelos Principales')
    normal(doc, (
        'Esta sección describe en detalle cada modelo ORM del módulo, sus campos, métodos y la '
        'lógica de negocio asociada. Todos los modelos residen en models/dgii_reports.py salvo '
        'que se indique lo contrario.'
    ))

    # 3.1 dgii.reports
    heading2(doc, '3.1 dgii.reports  —  dgii_reports.py')
    normal(doc, (
        'Modelo principal del módulo. Representa un período de declaración mensual (MM/YYYY) '
        'para una compañía. Contiene todos los datos agregados de los cuatro reportes y sirve '
        'como punto de entrada para la generación de archivos TXT y el formulario IT-1.'
    ))
    normal(doc, 'Campos de configuración:')
    add_table(doc,
        ['Campo', 'Tipo ORM', 'Descripción'],
        [
            ['name', 'Char(7), required, unique', 'Período en formato MM/YYYY. Ejemplo: 03/2025.'],
            ['state', 'Selection', 'Estado del reporte: draft, error, generated, sent.'],
            ['company_id', 'Many2one res.company', 'Compañía para la que se declara.'],
            ['currency_id', 'Many2one res.currency', 'Moneda del reporte (relacionada con company_id).'],
            ['start_date', 'Date (computed)', 'Primer día del período (1/MM/YYYY).'],
            ['end_date', 'Date (computed)', 'Último día del período (último del mes).'],
            ['previous_balance', 'Float', 'Saldo anterior a trasladar (ingreso manual).'],
            ['previous_report_pending', 'Boolean (computed)', 'True si existe un reporte anterior no enviado.'],
            ['dgii_consolidate_branches', 'Boolean (related)', 'Refleja la configuración de consolidación de la compañía.'],
        ],
        col_widths=[1.8, 1.8, 2.9]
    )
    normal(doc, 'Relaciones One2many (líneas de detalle por reporte):')
    add_table(doc,
        ['Campo', 'Modelo relacionado', 'Reporte'],
        [
            ['purchase_line_ids', 'dgii.reports.purchase.line', '606 — Compras y Gastos'],
            ['sale_line_ids', 'dgii.reports.sale.line', '607 — Ventas'],
            ['cancel_line_ids', 'dgii.reports.cancel.line', '608 — Anulados'],
            ['exterior_line_ids', 'dgii.reports.exterior.line', '609 — Pagos al Exterior'],
            ['it1_section_1_ids … it1_section_6_ids', 'dgii.reports.it1.line', 'Anexo A e IT-1'],
        ],
        col_widths=[2.2, 2.2, 2.1]
    )
    normal(doc, 'Campos de resumen 606 (computed desde purchase_line_ids):')
    add_table(doc,
        ['Campo', 'Descripción'],
        [
            ['purchase_records', 'Cantidad de líneas 606.'],
            ['service_total_amount', 'Total servicios sin ITBIS.'],
            ['good_total_amount', 'Total bienes sin ITBIS.'],
            ['purchase_invoiced_amount', 'Monto total facturado (servicios + bienes).'],
            ['purchase_invoiced_itbis', 'ITBIS facturado total.'],
            ['purchase_withholded_itbis', 'RITBIS retenido.'],
            ['cost_itbis', 'ITBIS llevado a costo.'],
            ['advance_itbis', 'ITBIS recuperable (crédito).'],
            ['income_withholding', 'ISR retenido total.'],
            ['purchase_selective_tax', 'ISC total.'],
            ['purchase_other_taxes', 'Otros impuestos.'],
            ['purchase_legal_tip', 'Propina legal.'],
            ['purchase_filename / purchase_binary', 'Nombre y contenido base64 del archivo TXT 606.'],
        ],
        col_widths=[2.4, 4.1]
    )
    normal(doc, 'Campos de resumen 607 (computed desde sale_line_ids):')
    add_table(doc,
        ['Campo', 'Descripción'],
        [
            ['sale_records', 'Cantidad de líneas 607.'],
            ['sale_invoiced_amount', 'Monto total facturado en ventas.'],
            ['sale_invoiced_itbis', 'ITBIS facturado en ventas.'],
            ['sale_withholded_itbis', 'RITBIS retenido por terceros.'],
            ['sale_withholded_isr', 'ISR retenido por terceros.'],
            ['sale_selective_tax', 'ISC en ventas.'],
            ['sale_other_taxes', 'Otros impuestos en ventas.'],
            ['sale_legal_tip', 'Propina legal en ventas.'],
            ['sale_filename / sale_binary', 'Nombre y contenido base64 del archivo TXT 607.'],
        ],
        col_widths=[2.4, 4.1]
    )
    normal(doc, 'Campos de resumen comprobantes consumidor (B02):')
    add_table(doc,
        ['Campo', 'Descripción'],
        [
            ['csmr_ncf_qty', 'Cantidad de NCF consumidor emitidos.'],
            ['csmr_ncf_total_amount', 'Monto total B02.'],
            ['csmr_ncf_total_itbis', 'ITBIS total B02.'],
            ['csmr_ncf_total_isc', 'ISC total B02.'],
            ['csmr_ncf_total_other', 'Otros impuestos B02.'],
            ['csmr_ncf_total_lgl_tip', 'Propina legal total B02.'],
            ['csmr_cash', 'Efectivo B02.'],
            ['csmr_bank', 'Cheque/Transferencia/Depósito B02.'],
            ['csmr_card', 'Tarjeta de crédito/débito B02.'],
            ['csmr_credit', 'Crédito a cuenta B02.'],
            ['csmr_bond', 'Bonos/Certificados B02.'],
            ['csmr_swap', 'Permuta B02.'],
            ['csmr_others', 'Otras formas de pago B02.'],
        ],
        col_widths=[2.4, 4.1]
    )
    normal(doc, 'Métodos clave:')
    add_table(doc,
        ['Método', 'Descripción'],
        [
            ['_compute_dates()', 'Calcula start_date y end_date a partir del campo name (MM/YYYY).'],
            ['_validate_date_format(date)', 'Valida que el nombre del período cumpla el formato MM/YYYY con meses 01-12.'],
            ['_get_companies_ids()', 'Retorna los IDs de la compañía actual y sus sucursales si consolidate=True.'],
            ['_get_invoices(states, types)', 'Obtiene facturas del período activo con los estados y tipos indicados.'],
            ['_get_pending_invoices(types, states)', 'Obtiene facturas de períodos anteriores cuya fecha de pago cae en el período actual.'],
            ['_compute_606_data()', 'Genera líneas purchase_line_ids y crea el TXT 606 en base64.'],
            ['_compute_607_data()', 'Genera líneas sale_line_ids, calcula resumen B02 y crea el TXT 607.'],
            ['_compute_608_data()', 'Genera líneas cancel_line_ids y crea el TXT 608.'],
            ['_compute_609_data()', 'Genera líneas exterior_line_ids con código país ISO y crea el TXT 609.'],
            ['_compute_attachment_a_and_it1_data()', 'Calcula y crea las líneas it1 (Anexo A + IT-1) a partir de las cuentas contables configuradas.'],
            ['_generate_report()', 'Ejecuta todos los _compute_*_data() secuencialmente, borra líneas previas.'],
            ['generate_report()', 'Punto de entrada desde el botón UI. Si ya fue generado, lanza el wizard de confirmación.'],
            ['state_sent()', 'Marca el reporte como enviado y llama a _invoice_status_sent() para marcar facturas.'],
            ['get_606_tree_view() … get_609_tree_view()', 'Devuelven acción para abrir la vista árbol del reporte correspondiente desde el form.'],
        ],
        col_widths=[2.8, 3.7]
    )

    # 3.2
    heading2(doc, '3.2 dgii.reports.purchase.line  —  dgii_reports.py')
    normal(doc, 'Detalle de cada factura de compra incluida en el reporte 606.')
    add_table(doc,
        ['Campo', 'Tipo', 'Descripción'],
        [
            ['dgii_report_id', 'Many2one dgii.reports', 'Reporte padre (cascade delete).'],
            ['rnc_cedula', 'Char', 'RNC o cédula del proveedor (11 o 9 dígitos).'],
            ['identification_type', 'Char', '1=RNC, 2=Cédula, 3=Pasaporte.'],
            ['l10n_do_expense_type', 'Char', 'Tipo de bienes o servicios (01-11).'],
            ['fiscal_invoice_number', 'Char', 'NCF de la factura (B01, B11, etc.).'],
            ['modified_invoice_number', 'Char', 'NCF modificado (notas de crédito/débito).'],
            ['invoice_date', 'Date', 'Fecha de emisión de la factura.'],
            ['payment_date', 'Date', 'Fecha de pago (última reconciliación).'],
            ['service_total_amount', 'Monetary', 'Monto servicios sin ITBIS.'],
            ['good_total_amount', 'Monetary', 'Monto bienes sin ITBIS.'],
            ['invoiced_amount', 'Monetary', 'Monto total facturado (servicios + bienes).'],
            ['invoiced_itbis', 'Monetary', 'ITBIS facturado.'],
            ['withholded_itbis', 'Monetary', 'RITBIS retenido.'],
            ['proportionality_tax', 'Monetary', 'ITBIS sujeto a proporcionalidad.'],
            ['cost_itbis', 'Monetary', 'ITBIS llevado a costo.'],
            ['advance_itbis', 'Monetary', 'ITBIS recuperable (crédito fiscal).'],
            ['isr_withholding_type', 'Char', 'Categoría de retención ISR (01-08).'],
            ['income_withholding', 'Monetary', 'ISR retenido.'],
            ['selective_tax', 'Monetary', 'ISC.'],
            ['other_taxes', 'Monetary', 'Otros impuestos.'],
            ['legal_tip', 'Monetary', 'Propina legal (10%).'],
            ['payment_type', 'Char', 'Forma de pago: 01-Efectivo … 07-Mixto.'],
            ['invoice_partner_id', 'Many2one res.partner', 'Proveedor de la factura.'],
            ['invoice_id', 'Many2one account.move', 'Factura de origen.'],
            ['company_id', 'Many2one res.company', 'Compañía (para consolidación multi-sucursal).'],
        ],
        col_widths=[1.9, 1.6, 3.0]
    )

    # 3.3
    heading2(doc, '3.3 dgii.reports.sale.line  —  dgii_reports.py')
    normal(doc, 'Detalle de cada factura de venta incluida en el reporte 607.')
    add_table(doc,
        ['Campo', 'Tipo', 'Descripción'],
        [
            ['dgii_report_id', 'Many2one dgii.reports', 'Reporte padre (cascade delete).'],
            ['rnc_cedula', 'Char', 'RNC o cédula del cliente.'],
            ['identification_type', 'Char', '1=RNC, 2=Cédula, 3=Pasaporte.'],
            ['fiscal_invoice_number', 'Char', 'NCF de la factura.'],
            ['modified_invoice_number', 'Char', 'NCF modificado (notas de crédito).'],
            ['income_type', 'Char', 'Tipo de ingreso (01-06).'],
            ['invoice_date', 'Date', 'Fecha de emisión.'],
            ['withholding_date', 'Date', 'Fecha de retención/pago.'],
            ['invoiced_amount', 'Monetary', 'Monto total facturado.'],
            ['invoiced_itbis', 'Monetary', 'ITBIS facturado.'],
            ['third_withheld_itbis', 'Monetary', 'RITBIS retenido por terceros.'],
            ['third_income_withholding', 'Monetary', 'ISR retenido por terceros.'],
            ['selective_tax', 'Monetary', 'ISC.'],
            ['other_taxes', 'Monetary', 'Otros impuestos.'],
            ['legal_tip', 'Monetary', 'Propina legal.'],
            ['cash', 'Monetary', 'Pago en efectivo.'],
            ['bank', 'Monetary', 'Pago por cheque/transferencia.'],
            ['card', 'Monetary', 'Pago con tarjeta.'],
            ['credit', 'Monetary', 'Crédito.'],
            ['bond', 'Monetary', 'Bono/Certificado.'],
            ['swap', 'Monetary', 'Permuta.'],
            ['others', 'Monetary', 'Otras formas de pago.'],
            ['invoice_partner_id', 'Many2one res.partner', 'Cliente de la factura.'],
            ['invoice_id', 'Many2one account.move', 'Factura de origen.'],
            ['company_id', 'Many2one res.company', 'Compañía.'],
        ],
        col_widths=[1.9, 1.6, 3.0]
    )

    # 3.4
    heading2(doc, '3.4 dgii.reports.cancel.line  —  dgii_reports.py')
    normal(doc, 'Registro de comprobantes anulados incluidos en el reporte 608.')
    add_table(doc,
        ['Campo', 'Tipo', 'Descripción'],
        [
            ['dgii_report_id', 'Many2one dgii.reports', 'Reporte padre (cascade delete).'],
            ['fiscal_invoice_number', 'Char', 'NCF anulado.'],
            ['invoice_date', 'Date', 'Fecha de la factura anulada.'],
            ['annulation_type', 'Char', 'Tipo de anulación DGII.'],
            ['invoice_partner_id', 'Many2one res.partner', 'Contacto relacionado.'],
            ['invoice_id', 'Many2one account.move', 'Factura de origen.'],
            ['company_id', 'Many2one res.company', 'Compañía.'],
        ],
        col_widths=[1.9, 1.6, 3.0]
    )

    # 3.5
    heading2(doc, '3.5 dgii.reports.exterior.line  —  dgii_reports.py')
    normal(doc, 'Detalle de pagos al exterior incluidos en el reporte 609 (facturas B17).')
    add_table(doc,
        ['Campo', 'Tipo', 'Descripción'],
        [
            ['dgii_report_id', 'Many2one dgii.reports', 'Reporte padre (cascade delete).'],
            ['legal_name', 'Char', 'Razón social del proveedor extranjero.'],
            ['tax_id_type', 'Char', '1=RNC, 2=Cédula, 3=Pasaporte, 4=NIF exterior.'],
            ['tax_id', 'Char', 'Número de identificación fiscal del proveedor.'],
            ['country_code', 'Char', 'Código numérico ISO 3166 del país del proveedor.'],
            ['purchased_service_type', 'Char', 'Categoría de servicio adquirido (01-08).'],
            ['service_type_detail', 'Many2one invoice.service.type.detail', 'Subcategoría detallada del servicio.'],
            ['related_part', 'Char', '0=No relacionado, 1=Relacionado (parte vinculada).'],
            ['doc_number', 'Char', 'Número de documento exterior.'],
            ['doc_date', 'Date', 'Fecha del documento.'],
            ['invoiced_amount', 'Monetary', 'Monto facturado.'],
            ['presumed_income', 'Monetary', 'Renta presunta (cálculo pendiente de implementación).'],
            ['withholded_isr', 'Monetary', 'ISR retenido al proveedor exterior.'],
            ['invoice_id', 'Many2one account.move', 'Factura de origen.'],
            ['company_id', 'Many2one res.company', 'Compañía.'],
        ],
        col_widths=[1.9, 1.6, 3.0]
    )

    # 3.6
    heading2(doc, '3.6 dgii.reports.it1.line  —  dgii_reports.py')
    normal(doc, (
        'Líneas de las secciones del Anexo A e IT-1. Cada registro corresponde a una casilla '
        'o título de sección en los formularios fiscales.'
    ))
    add_table(doc,
        ['Campo', 'Tipo', 'Descripción'],
        [
            ['dgii_report_id', 'Many2one dgii.reports', 'Reporte padre (cascade delete).'],
            ['name', 'Char', 'Descripción de la casilla o sección.'],
            ['sequence', 'Integer', 'Orden de aparición en la sección.'],
            ['section', 'Selection (1-6)', 'Sección del formulario (1=II…6=VI del Anexo A; secciones IT-1).'],
            ['display_type', 'Selection', 'line_section (encabezado) o line_note (casilla de valor).'],
            ['coefficient', 'Float', 'Coeficiente porcentual aplicado (e.g. 18% ITBIS).'],
            ['quantity', 'Integer', 'Cantidad de operaciones.'],
            ['local_purchase', 'Monetary', 'Monto de compras locales.'],
            ['services', 'Monetary', 'Monto de servicios.'],
            ['imports', 'Monetary', 'Monto de importaciones.'],
            ['amount', 'Monetary', 'Valor total de la casilla.'],
            ['currency_id', 'Many2one res.currency (related)', 'Moneda del reporte padre.'],
        ],
        col_widths=[1.9, 1.6, 3.0]
    )

    # 3.7
    heading2(doc, '3.7 invoice.service.type.detail  —  invoice_service_type_detail.py')
    normal(doc, (
        'Catálogo jerárquico de tipos de servicio para facturas de proveedores del exterior (B17). '
        'Tiene dos niveles: categoría padre (01-08) y subcategorías con código de dos dígitos.'
    ))
    add_table(doc,
        ['Campo', 'Tipo', 'Descripción'],
        [
            ['name', 'Char', 'Descripción del tipo o subtipo de servicio.'],
            ['code', 'Char(2), unique', 'Código de dos caracteres (e.g. 11, 21, 31).'],
            ['parent_code', 'Char', 'Código de la categoría padre (01-08). Nulo en categorías raíz.'],
        ],
        col_widths=[1.5, 1.5, 3.5]
    )
    normal(doc, 'Categorías principales y sus subcategorías:')
    add_table(doc,
        ['Código', 'Categoría', 'Subcategorías'],
        [
            ['01', 'Personal', '11-Sueldos/Salarios, 12-Otros gastos de personal'],
            ['02', 'Trabajo/Suministros/Servicios', '21-Servicios profesionales (empresas), 22-Serv. prof. (personas), 23-Seguridad/mensajería/transporte (personas), 24-Seguridad/mensajería (empresas)'],
            ['03', 'Arrendamientos', '31-Bienes inmuebles (personas), 32-Bienes inmuebles (empresas), 33-Otros arrendamientos'],
            ['04', 'Activos Fijos', '41-Reparaciones, 42-Mantenimiento'],
            ['05', 'Representación', '51-Relaciones públicas, 52-Pub./Materiales promocionales, 53-Promocionales, 54-Otras representaciones'],
            ['06', 'Financieros', '61-Intereses préstamos, 62-Intereses financiamiento'],
            ['07', 'Seguros', '(Subtypes por configurar según normativa)'],
            ['08', 'Regalías/Intangibles', '(Subtypes por configurar según normativa)'],
        ],
        col_widths=[0.6, 1.8, 4.1]
    )

    doc.add_page_break()

    # ── Section 4 ─────────────────────────────────────────────────────────────
    heading1(doc, '4. Extensiones a Modelos Base')
    normal(doc, (
        'El módulo extiende varios modelos base de Odoo para agregar los campos necesarios '
        'para la generación de las declaraciones. Todas las extensiones se realizan mediante '
        'herencia ORM (_inherit).'
    ))

    # 4.1 account.move
    heading2(doc, '4.1 account.move  —  account_move.py')
    normal(doc, (
        'Agrega campos de cómputo fiscal a las facturas. Todos los campos son computed y '
        'stored=True para permitir búsquedas eficientes en los reportes.'
    ))
    normal(doc, 'Campos de impuestos (computed, stored, monetary):')
    add_table(doc,
        ['Campo', 'Tipo impuesto', 'Descripción'],
        [
            ['invoiced_itbis', 'itbis', 'ITBIS facturado (líneas de impuesto clasificadas como itbis).'],
            ['proportionality_tax', 'itbis (proporcional)', 'ITBIS sujeto a regla de proporcionalidad.'],
            ['cost_itbis', 'itbis (costo)', 'ITBIS que va a costo (no deducible).'],
            ['advance_itbis', 'itbis (crédito)', 'ITBIS recuperable como crédito fiscal.'],
            ['withholding_itbis', 'ritbis', 'RITBIS retenido (tax_type=ritbis).'],
            ['income_withholding', 'isr', 'ISR retenido (tax_type=isr).'],
            ['selective_tax', 'isc', 'Impuesto selectivo al consumo.'],
            ['other_taxes', 'other', 'Otros impuestos no clasificados.'],
            ['legal_tip', 'tip', 'Propina legal (10%).'],
        ],
        col_widths=[1.8, 1.6, 3.1]
    )
    normal(doc, 'Campos de montos por categoría de producto:')
    add_table(doc,
        ['Campo', 'Descripción'],
        [
            ['service_total_amount', 'Subtotal de líneas de servicios (producto tipo service) sin ITBIS.'],
            ['good_total_amount', 'Subtotal de líneas de bienes (producto tipo consu/storable) sin ITBIS.'],
        ],
        col_widths=[2.0, 4.5]
    )
    normal(doc, 'Campos específicos de compras (in_invoice / in_refund):')
    add_table(doc,
        ['Campo', 'Descripción'],
        [
            ['isr_withholding_type', 'Código de categoría ISR (01-08) extraído del impuesto de retención.'],
            ['payment_form', 'Forma de pago: 01-Efectivo, 02-Cheque/Transf., 03-Tarjeta, 04-Crédito, 05-Permuta, 06-Nota de crédito, 07-Mixto.'],
            ['payment_date', 'Fecha del último pago (extraída del widget de pagos, stored=True).'],
        ],
        col_widths=[2.0, 4.5]
    )
    normal(doc, 'Campos específicos de facturas de exterior (B17):')
    add_table(doc,
        ['Campo', 'Descripción'],
        [
            ['service_type', 'Categoría de servicio adquirido al proveedor exterior (01-08).'],
            ['service_type_detail', 'Many2one a invoice.service.type.detail, filtrado por service_type.'],
            ['is_exterior', 'Boolean computed: True si el NCF comienza con "B17".'],
        ],
        col_widths=[2.0, 4.5]
    )
    normal(doc, 'Campo de estado fiscal:')
    add_table(doc,
        ['Valor', 'Etiqueta', 'Significado'],
        [
            ['normal', 'Parcial', 'Factura en período activo, no reportada aún.'],
            ['done', 'Reportada', 'Factura incluida en un reporte marcado como enviado.'],
            ['blocked', 'No enviada', 'Factura bloqueada manualmente de ser reportada.'],
        ],
        col_widths=[1.2, 1.5, 3.8]
    )
    normal(doc, 'Constraint:')
    bullet(doc, '_check_isr_tax(): ValidationError si se configuran más de un impuesto ISR o RITBIS en la misma factura.')
    normal(doc, 'Métodos helpers clave:')
    add_table(doc,
        ['Método', 'Descripción'],
        [
            ['_get_invoice_payment_widget()', 'Parsea el JSON del widget de pagos de Odoo.'],
            ['_compute_invoice_payment_date()', 'Extrae la fecha del último pago registrado.'],
            ['_convert_to_local_currency(amount)', 'Convierte montos a la moneda de la compañía con manejo de signos.'],
            ['_compute_amount_fields()', 'Divide las líneas de factura en servicios y bienes.'],
            ['_compute_taxes_fields()', 'Agrega montos de impuesto clasificados por l10n_do_tax_type.'],
            ['_compute_withholding_taxes()', 'Extrae RITBIS e ISR de las líneas de impuesto.'],
            ['_compute_isr_withholding_type()', 'Lee la categoría ISR del impuesto de retención configurado.'],
            ['_compute_in_invoice_payment_form()', 'Determina el código de forma de pago a partir del tipo de diario.'],
            ['_compute_is_exterior()', 'Detecta si la factura es de proveedor exterior por prefijo NCF B17.'],
            ['_get_payment_string()', 'Analiza las reconciliaciones para determinar la forma de pago.'],
            ['_get_sale_payments_forms(invoice_id)', 'Calcula el desglose de formas de pago para el reporte 607.'],
            ['norma_recompute()', 'Fuerza el recómputo de todos los campos DGII stored en las facturas seleccionadas.'],
        ],
        col_widths=[2.8, 3.7]
    )

    # 4.2 account.tax
    heading2(doc, '4.2 account.tax  —  account_tax.py')
    normal(doc, 'Agrega clasificación DGII a los impuestos contables.')
    add_table(doc,
        ['Campo', 'Tipo', 'Valores', 'Descripción'],
        [
            ['l10n_do_tax_type', 'Selection', 'itbis, ritbis, isr, isc, other, tip, rext, none', 'Tipo de impuesto para clasificación DGII. Utilizado por account_move para calcular campos computados.'],
            ['isr_retention_type', 'Selection', '01-Arrendamientos, 02-Honorarios, 03-Otros ingresos, 04-Presuntos, 05-Intereses corp., 06-Intereses personas, 07-Proveedores estado, 08-Juegos móviles', 'Categoría de retención ISR para el campo isr_withholding_type de la factura.'],
        ],
        col_widths=[1.5, 0.9, 2.1, 2.0]
    )

    # 4.3 account.journal
    heading2(doc, '4.3 account.journal  —  account_journal.py')
    normal(doc, 'Agrega la forma de pago predeterminada al diario contable.')
    add_table(doc,
        ['Campo', 'Tipo', 'Descripción'],
        [
            ['l10n_do_fiscal_journal', 'Boolean', 'Indica si el diario es un diario fiscal (para filtros de reportes).'],
            ['payment_form', 'Selection', 'Forma de pago por defecto: cash, bank (cheque/transf.), card, credit, swap, bond, others.'],
        ],
        col_widths=[2.0, 1.2, 3.3]
    )

    # 4.4 account.account
    heading2(doc, '4.4 account.account  —  account_account.py')
    normal(doc, (
        'Agrega la asignación de cuentas contables a casillas del Anexo A e IT-1. '
        'Al configurar una cuenta con una casilla, los saldos de esa cuenta se suman '
        'automáticamente a la casilla correspondiente al generar el reporte.'
    ))
    add_table(doc,
        ['Campo', 'Descripción'],
        [
            ['box_attachment_a', 'Selection con 56 opciones (A9, A10, A27-A32, A34-A36, A39-A40, A45-A47, A49-A51, A53). Asigna la cuenta a una casilla del Anexo A.'],
            ['box_it1', 'Selection con 68 opciones (I2-I8, I15, I28, I31-I32, I35-I37, I39-I40, I42-I45, I47-I48, I59, I61, I64-I66). Asigna la cuenta a una casilla del IT-1.'],
        ],
        col_widths=[1.8, 4.7]
    )

    # 4.5 res.company
    heading2(doc, '4.5 res.company  —  res_company.py')
    normal(doc, 'Agrega soporte para consolidación multi-sucursal en los reportes DGII.')
    add_table(doc,
        ['Campo', 'Descripción'],
        [
            ['dgii_consolidate_branches', 'Boolean. Cuando True en la compañía padre, el reporte incluye datos de todas las compañías hijas. Solo válido en compañías raíz (constraint _check_consolidate_branches).'],
        ],
        col_widths=[2.2, 4.3]
    )
    normal(doc, 'Constraint: _check_consolidate_branches() — ValidationError si se activa en una compañía que no es raíz (tiene parent_id).')

    # 4.6 res.partner
    heading2(doc, '4.6 res.partner  —  res_partner.py (extensión dgii_reports)')
    normal(doc, 'Agrega clasificación de partes relacionadas para el reporte 609.')
    add_table(doc,
        ['Campo', 'Tipo', 'Descripción'],
        [
            ['related', 'Selection', '0=No relacionado (default), 1=Relacionado. Utilizado en el reporte 609 para identificar partes vinculadas (Column related_part).'],
        ],
        col_widths=[1.5, 1.2, 3.8]
    )

    doc.add_page_break()

    # ── Section 5 ─────────────────────────────────────────────────────────────
    heading1(doc, '5. Wizards (Modelos Transitorios)')
    heading2(doc, '5.1 dgii.report.regenerate.wizard')
    normal(doc, (
        'Diálogo de confirmación que se presenta al usuario cuando intenta regenerar un reporte '
        'que ya ha sido generado previamente. Protege contra la regeneración accidental, ya que '
        'la acción borra todas las líneas existentes y recalcula desde cero.'
    ))
    add_table(doc,
        ['Campo', 'Tipo', 'Descripción'],
        [
            ['report_id', 'Many2one dgii.reports', 'Referencia al reporte que se desea regenerar.'],
        ],
        col_widths=[1.5, 1.8, 3.2]
    )
    add_table(doc,
        ['Método', 'Descripción'],
        [
            ['regenerate()', 'Llama a _generate_report() en el reporte referenciado y cierra el wizard.'],
        ],
        col_widths=[1.8, 4.7]
    )

    doc.add_page_break()

    # ── Section 6 ─────────────────────────────────────────────────────────────
    heading1(doc, '6. Controladores HTTP')
    heading2(doc, '6.1 DgiiReportsControllers  —  controllers/main.py')
    normal(doc, (
        'Controlador web que provee URLs directas para navegar a facturas o contactos '
        'desde reportes externos o herramientas de auditoría.'
    ))
    add_table(doc,
        ['Atributo', 'Valor'],
        [
            ['Ruta', '/dgii_reports/<ncf_rnc>'],
            ['Método HTTP', 'GET'],
            ['Autenticación', 'user (sesión activa requerida)'],
        ],
        col_widths=[1.8, 4.7]
    )
    normal(doc, 'Lógica:')
    bullet(doc, 'Si ncf_rnc empieza con "B": redirige a la factura con fiscal_invoice_number == ncf_rnc.')
    bullet(doc, 'Si no: redirige al contacto con vat == ncf_rnc.')
    bullet(doc, 'Retorna URL de formulario Odoo con acción y contexto correspondiente.')

    doc.add_page_break()

    # ── Section 7 ─────────────────────────────────────────────────────────────
    heading1(doc, '7. Seguridad y Control de Acceso')
    heading2(doc, '7.1 Permisos de Acceso por Modelo')
    normal(doc, 'Todos los modelos del módulo otorgan permisos completos al grupo account.group_account_user:')
    add_table(doc,
        ['Modelo', 'Leer', 'Escribir', 'Crear', 'Eliminar'],
        [
            ['dgii.reports', '✓', '✓', '✓', '✓'],
            ['dgii.reports.purchase.line', '✓', '✓', '✓', '✓'],
            ['dgii.reports.sale.line', '✓', '✓', '✓', '✓'],
            ['dgii.reports.cancel.line', '✓', '✓', '✓', '✓'],
            ['dgii.reports.exterior.line', '✓', '✓', '✓', '✓'],
            ['dgii.reports.it1.line', '✓', '✓', '✓', '✓'],
            ['invoice.service.type.detail', '✓', '✓', '✓', '✓'],
            ['dgii.report.regenerate.wizard', '✓', '✓', '✓', '✓'],
        ],
        col_widths=[2.6, 0.7, 0.7, 0.7, 0.7]
    )
    normal(doc, (
        'NOTA: No existen reglas de registro (ir.rule) adicionales. El acceso se controla únicamente '
        'a nivel de modelo mediante el grupo de contabilidad estándar de Odoo.'
    ))

    doc.add_page_break()

    # ── Section 8 ─────────────────────────────────────────────────────────────
    heading1(doc, '8. Datos Maestros')
    normal(doc, (
        'El módulo carga datos iniciales al instalar mediante el archivo '
        'data/invoice_service_type_detail_data.xml. Este archivo define el catálogo completo '
        'de tipos y subtipos de servicio para el reporte 609 (B17 — Pagos al Exterior).'
    ))
    add_table(doc,
        ['ID XML', 'Código', 'Descripción'],
        [
            ['service_type_01', '01', 'Personal'],
            ['service_type_11', '11', 'Sueldos, salarios y remuneraciones'],
            ['service_type_12', '12', 'Otros gastos de personal'],
            ['service_type_02', '02', 'Trabajo, suministros y servicios'],
            ['service_type_21', '21', 'Servicios profesionales prestados por personas jurídicas'],
            ['service_type_22', '22', 'Servicios profesionales prestados por personas físicas'],
            ['service_type_23', '23', 'Seguridad, mensajería y transporte (personas físicas)'],
            ['service_type_24', '24', 'Seguridad, mensajería y transporte (personas jurídicas)'],
            ['service_type_03', '03', 'Arrendamientos'],
            ['service_type_31', '31', 'Arrendamiento de bienes inmuebles (personas físicas)'],
            ['service_type_32', '32', 'Arrendamiento de bienes inmuebles (personas jurídicas)'],
            ['service_type_33', '33', 'Otros arrendamientos'],
            ['service_type_04', '04', 'Activos fijos'],
            ['service_type_41', '41', 'Reparaciones a activos fijos'],
            ['service_type_42', '42', 'Mantenimiento a activos fijos'],
            ['service_type_05', '05', 'Representación'],
            ['service_type_51', '51', 'Relaciones públicas'],
            ['service_type_52', '52', 'Publicidad y materiales promocionales'],
            ['service_type_53', '53', 'Materiales promocionales'],
            ['service_type_54', '54', 'Otras representaciones'],
            ['service_type_06', '06', 'Financieros'],
            ['service_type_61', '61', 'Intereses de préstamos'],
            ['service_type_62', '62', 'Intereses de financiamiento'],
        ],
        col_widths=[1.8, 0.8, 3.9]
    )

    doc.add_page_break()

    # ── Section 9 ─────────────────────────────────────────────────────────────
    heading1(doc, '9. Hooks de Instalación  —  hooks.py')
    normal(doc, (
        'El archivo hooks.py define post_init_hook(env), que se ejecuta automáticamente '
        'al completar la instalación del módulo. Realiza las siguientes configuraciones iniciales:'
    ))
    heading2(doc, '9.1 set_l10n_do_tax_types(cr)')
    normal(doc, (
        'Actualiza el campo l10n_do_tax_type en los impuestos existentes, basándose en el '
        'nombre técnico de la plantilla de impuesto (tax template) de l10n_do. '
        'Opera directamente en SQL para mayor eficiencia.'
    ))
    add_table(doc,
        ['Tipo DGII', 'Plantillas asignadas (ejemplos)'],
        [
            ['itbis', 'tax_18_sale, tax_0_sale, tax_18_of_10, tax_tip_sale, tax_18_purch, tax_16_purch, tax_9_purch, tax_8_purch, tax_18_purch_serv, tax_18_importation, tax_0_purch, tax_tip_purch'],
            ['ritbis', 'ret_100_tax_security, ret_100_tax_nonprofit, ret_100_tax_person, ret_30_tax_moral, ret_30_tax_freelance, ret_75_tax_nonformal'],
            ['isr', 'ret_10_income_person, ret_10_income_rent, ret_10_income_dividend, ret_2_income_person, ret_2_income_transfer, ret_27_income_remittance, ret_5_income_gov'],
            ['isc', 'tax_10_telco'],
            ['other', 'tax_2_telco, tax_0015_bank'],
        ],
        col_widths=[1.2, 5.3]
    )
    normal(doc, 'También configura isr_retention_type por plantilla:')
    add_table(doc,
        ['Categoría ISR', 'Plantillas'],
        [
            ['01 — Arrendamientos', 'ret_10_income_rent'],
            ['02 — Honorarios de servicios', 'ret_10_income_person, ret_2_income_person, ret_2_income_transfer'],
            ['03 — Otros ingresos', 'ret_10_income_dividend, ret_27_income_remittance'],
            ['07 — Proveedores del estado', 'ret_5_income_gov'],
        ],
        col_widths=[2.2, 4.3]
    )

    heading2(doc, '9.2 set_journal_payment_forms(cr)')
    normal(doc, (
        'Establece la forma de pago predeterminada en diarios de efectivo y banco '
        'que aún no tienen una asignada (payment_form IS NULL):'
    ))
    add_table(doc,
        ['Tipo de diario', 'payment_form asignado'],
        [
            ['cash', 'cash (01 — Efectivo)'],
            ['bank', 'bank (02 — Cheque/Transferencia/Depósito)'],
        ],
        col_widths=[2.5, 4.0]
    )

    heading2(doc, '9.3 Datos de Demostración  —  demo/demo_dgii.xml')
    normal(doc, (
        'El módulo incluye datos de demostración comprensivos que cubren la totalidad de los campos '
        'y columnas de los cuatro reportes DGII. Al instalar con datos demo activados, se ejecuta '
        'automáticamente la siguiente secuencia a través del archivo demo/demo_dgii.xml, '
        'que invoca dgii.reports._load_demo_data().'
    ))
    normal(doc, 'Secuencia de ejecución al instalar con demo data:')
    add_table(doc,
        ['Orden', 'Método', 'Qué hace'],
        [
            ['1', 'l10n_do_accounting\n_post_load_demo_data()', 'Crea la compañía demo, carga el plan de cuentas DO y genera los NCF Batches necesarios para todos los tipos de comprobante.'],
            ['2', 'dgii_reports\n_setup_demo_company_config()', 'Configura l10n_do_tax_type en todos los impuestos del chart. Establece isr_retention_type. Asigna payment_form a diarios cash/bank. Crea diario de Tarjeta (CARD).'],
            ['3', 'dgii_reports\n_create_demo_regular_invoices()', 'Crea y confirma facturas comprensivas para 606, 607 y 608. Ver detalle abajo.'],
            ['4', 'dgii_reports\n_create_demo_exterior_invoices()', 'Crea facturas B17 para el reporte 609 cubriendo todos los tipos de servicio y casos de ISR.'],
            ['5', 'dgii_reports\n_load_demo_data()', 'Genera el reporte DGII para el mes anterior, consolidando todo lo anterior.'],
        ],
        col_widths=[0.6, 2.2, 3.7]
    )

    heading3(doc, '_setup_demo_company_config(company)')
    normal(doc, 'Configura los impuestos de la compañía demo con los tipos correctos para los reportes:')
    add_table(doc,
        ['l10n_do_tax_type', 'Impuestos configurados'],
        [
            ['itbis', 'tax_18_sale, tax_0_sale, tax_18_purch, tax_18_purch_serv'],
            ['ritbis', 'ret_100_tax_person, ret_75_tax_nonformal, ret_30_tax_moral, ret_100_tax_nonprofit, ret_100_tax_security'],
            ['isr', 'ret_10_income_person (tipo 02), ret_5_income_gov (tipo 07), ret_27_income_remittance (tipo 03)'],
            ['tip', 'tax_tip_sale'],
            ['isc', 'tax_10_telco'],
        ],
        col_widths=[1.4, 5.1]
    )
    normal(doc, (
        'Además crea un diario de tipo cash con código "CARD" y l10n_do_payment_form="card" '
        'para cubrir la columna "Tarjeta" del reporte 607.'
    ))

    heading3(doc, '_create_demo_regular_invoices(company) — Cobertura por reporte')
    normal(doc, 'Crea dos productos: "Servicio Demo DGII" (type=service) y "Bien/Mercancía Demo DGII" (type=consu) para que los campos service_total_amount y good_total_amount del 606 tengan valores.')
    normal(doc, 'Facturas de ventas (607) generadas:')
    add_table(doc,
        ['Tipo NCF', 'Income Type', 'Impuestos', 'Pago', 'Columnas 607 cubiertas'],
        [
            ['B01', '01', 'tax_18_sale', 'Efectivo', 'invoiced_itbis, cash'],
            ['B01', '02', 'tax_18_sale', 'Banco', 'bank'],
            ['B01', '03', 'tax_18_sale', 'Banco', 'income_type 03'],
            ['B01', '04, 05, 06', 'tax_18_sale', 'Sin pago', 'credit, tipos de ingreso'],
            ['B01', '01', 'tax_18_sale + ret_30_tax_moral', 'Sin pago', 'third_withheld_itbis'],
            ['B15', '01', 'tax_18_sale + ret_5_income_gov', 'Sin pago', 'third_income_withholding'],
            ['B01', '01', 'tax_18_sale + tax_10_telco', 'Banco', 'selective_tax'],
            ['B02 (<250k)', '01', 'tax_18_sale', 'Efectivo', 'csmr_* (resumen consumidor)'],
            ['B02 (≥250k)', '01', 'tax_18_sale', 'Banco', 'TXT 607 incluido'],
            ['B02', '01', 'tax_18_sale + tax_tip_sale', 'Tarjeta', 'legal_tip, card'],
            ['B14', '01', 'tax_0_sale', 'Sin pago', 'ITBIS=0, zona franca'],
            ['B16', '01', 'tax_0_sale', 'Banco', 'exportación'],
            ['B04 (out_refund)', '01', 'tax_18_sale', '—', 'modified_invoice_number, credit_note'],
            ['B03 (debit note)', '02', 'tax_18_sale', '—', 'modified_invoice_number'],
        ],
        col_widths=[0.9, 0.9, 2.0, 0.9, 2.1]
    )
    normal(doc, 'Facturas de compras (606) generadas — un caso por cada tipo de gasto (01-11):')
    add_table(doc,
        ['Expense Type', 'Descripción', 'Impuestos', 'Producto', 'Pago', 'Columnas 606 cubiertas'],
        [
            ['01', 'Servicios de limpieza', 'tax_18_purch', 'Servicio', 'Banco', 'service_total_amount, payment_type 02'],
            ['02', 'Insumos de trabajo', 'tax_18_purch', 'Bien', 'Efectivo', 'good_total_amount, payment_type 01'],
            ['02', 'Honorarios + mat.', 'tax_18_purch + rit100 + isr10', 'Serv+Bien', 'Sin pago', 'withholded_itbis, income_withholding, isr_type 02'],
            ['03', 'Alquiler local', 'tax_18_purch', 'Servicio', 'Banco', 'expense_type 03'],
            ['04', 'Equipo de cómputo', 'tax_18_purch', 'Bien', 'Tarjeta', 'payment_type 03 (card)'],
            ['05', 'Gastos representación', 'tax_18_purch', 'Servicio', 'Sin pago', 'payment_type 04 (credit)'],
            ['06', 'Construcción', 'tax_18_purch + rit75', 'Servicio', 'Banco', 'withholded_itbis (75%)'],
            ['07', 'Intereses bancarios', 'tax_18_purch', 'Servicio', 'Sin pago', 'expense_type 07'],
            ['08', 'Gasto extraordinario', 'tax_18_purch', 'Servicio', 'Sin pago', 'expense_type 08'],
            ['09', 'Costo de ventas', 'tax_18_purch', 'Bien', 'Efectivo', 'expense_type 09'],
            ['10', 'Activo intangible', 'tax_18_purch', 'Bien', 'Banco', 'expense_type 10'],
            ['11', 'Seguro corporativo', 'tax_18_purch', 'Servicio', 'Sin pago', 'expense_type 11'],
            ['02', 'Telecomunicaciones', 'tax_18_purch + tax_10_telco', 'Servicio', 'Banco', 'selective_tax (ISC)'],
            ['02', 'Nota crédito compra', 'tax_18_purch', 'Servicio', '—', 'modified_invoice_number, credit_note'],
        ],
        col_widths=[0.7, 1.5, 1.6, 0.7, 0.7, 2.5]
    )
    normal(doc, 'Comprobantes anulados (608): se crean 10 facturas de venta y se cancelan con los tipos de anulación 01 al 10.')

    heading3(doc, '_create_demo_exterior_invoices(company) — Cobertura 609')
    add_table(doc,
        ['service_type', 'Partner', 'related_part', 'Con ISR 27%', 'Pagado', 'Columnas 609 cubiertas'],
        [
            ['01', 'España (empresa)', '0', 'No', 'Efectivo', 'service_type 01, tax_id_type 2'],
            ['02', 'España (empresa)', '0', 'Sí', 'Banco', 'withholded_isr, isr_withholding_date'],
            ['02', 'USA (empresa)', '0', 'No', 'Efectivo', 'country_code US'],
            ['03', 'USA (empresa)', '0', 'Sí', 'Banco', 'withholded_isr, service_type 03'],
            ['03', 'Colombia (empresa)', '1', 'No', 'Sin pago', 'related_part=1, isr_date vacío'],
            ['04', 'España', '0', 'No', 'Sin pago', 'service_type 04'],
            ['05', 'USA', '0', 'Sí', 'Banco', 'service_type 05, withholded_isr'],
            ['06', 'España', '0', 'No', 'Efectivo', 'service_type 06'],
            ['07', 'Colombia', '0', 'No', 'Sin pago', 'service_type 07'],
            ['08', 'USA', '1', 'Sí', 'Banco', 'service_type 08, related_part=1'],
        ],
        col_widths=[0.9, 1.3, 0.9, 0.9, 0.9, 2.6]
    )

    doc.add_page_break()

    # ── Section 10 ─────────────────────────────────────────────────────────────
    heading1(doc, '10. Reglas de Negocio Críticas')

    heading2(doc, '10.1 Flujo de Generación de Reportes')
    normal(doc, 'El proceso completo de generación sigue el siguiente orden:')
    add_table(doc,
        ['Paso', 'Acción', 'Método'],
        [
            ['1', 'Usuario crea reporte con período MM/YYYY', 'create() / ORM'],
            ['2', 'Usuario hace clic en "Generar Declaraciones"', 'generate_report()'],
            ['3', 'Si ya fue generado → muestra wizard de confirmación', 'dgii.report.regenerate.wizard'],
            ['4', 'Borrar líneas anteriores de todos los reportes', '_generate_report() → unlink()'],
            ['5', 'Generar líneas 606 y TXT', '_compute_606_data()'],
            ['6', 'Generar líneas 607, resumen B02 y TXT', '_compute_607_data()'],
            ['7', 'Generar líneas 608 y TXT', '_compute_608_data()'],
            ['8', 'Generar líneas 609 y TXT', '_compute_609_data()'],
            ['9', 'Calcular Anexo A e IT-1', '_compute_attachment_a_and_it1_data()'],
            ['10', 'Estado → generated', '_generate_report()'],
            ['11', 'Usuario descarga TXTs y sube a portal DGII', 'Manual (fuera de Odoo)'],
            ['12', 'Usuario hace clic en "Marcar como enviado"', 'state_sent()'],
            ['13', 'Facturas del período → fiscal_status = done', '_invoice_status_sent()'],
        ],
        col_widths=[0.4, 3.4, 2.7]
    )

    heading2(doc, '10.2 Inclusión de Facturas en Reportes')
    add_table(doc,
        ['Tipo de factura', 'Criterio de inclusión', 'Reporte'],
        [
            ['Compras activas del período', 'type in (in_invoice, in_refund) AND state=posted AND invoice_date in período AND fiscal_status != blocked', '606'],
            ['Compras pagadas de períodos anteriores', 'fiscal_status=normal AND payment_state in (paid, in_payment) AND payment_date in período actual', '606'],
            ['Ventas activas del período', 'type in (out_invoice, out_refund) AND state=posted AND invoice_date in período AND fiscal_status != blocked', '607'],
            ['Ventas pagadas de períodos anteriores', 'Igual que compras pendientes, para out_invoice', '607'],
            ['Anuladas del período', 'type in (out_invoice, in_invoice, out_refund) AND state=cancel AND invoice_date in período', '608'],
            ['Exterior del período', 'in_invoice AND state=posted AND ncf starts with B17 AND invoice_date in período', '609'],
        ],
        col_widths=[1.8, 3.0, 0.7]
    )

    heading2(doc, '10.3 Manejo Especial de Tipos de Comprobante')
    add_table(doc,
        ['Tipo NCF', 'Comportamiento especial'],
        [
            ['B02 — Consumidor', 'En 607: se cuenta en el resumen de consumidor pero se EXCLUYE del TXT si el monto < RD$250,000. Se incluye en el TXT si monto ≥ RD$250,000.'],
            ['B17 — Exterior', 'En 606: se usa el RNC de la propia compañía en lugar del RNC del proveedor. Se incluye también en 609.'],
            ['B12 — Consumo interno', 'En 607: se usa el RNC del cliente en lugar de blanco.'],
            ['Notas de crédito', 'Incluyen el NCF original en modified_invoice_number. En 607, las formas de pago se registran con montos negativos.'],
        ],
        col_widths=[1.6, 4.9]
    )

    heading2(doc, '10.4 Consolidación Multi-Sucursal')
    normal(doc, (
        'Cuando dgii_consolidate_branches = True en la compañía padre, el método '
        '_get_companies_ids() retorna los IDs de la compañía padre y todas sus '
        'compañías hijas (recursivo). Las consultas de facturas se ejecutan con el '
        'contexto allowed_company_ids incluyendo todas las compañías del grupo.'
    ))
    bullet(doc, 'Solo válido en compañías raíz (sin parent_id). Constraint _check_consolidate_branches() lo valida.')
    bullet(doc, 'Las compañías sucursal muestran un mensaje informativo en su formulario en lugar del toggle.')
    bullet(doc, 'Cada línea de reporte incluye company_id para identificar la sucursal de origen.')

    heading2(doc, '10.5 Restricciones de Integridad')
    add_table(doc,
        ['Restricción', 'Modelo', 'Error'],
        [
            ['Período único por compañía', 'dgii.reports', 'ValidationError si ya existe un reporte con el mismo name y company_id.'],
            ['Formato de período', 'dgii.reports', 'ValidationError si el campo name no cumple MM/YYYY con mes 01-12.'],
            ['Un solo ISR/RITBIS por factura', 'account.move', 'ValidationError (_check_isr_tax) si hay más de un impuesto de tipo isr o ritbis.'],
            ['Consolidación solo en raíz', 'res.company', 'ValidationError si se activa dgii_consolidate_branches en una empresa con parent_id.'],
        ],
        col_widths=[1.8, 1.5, 3.2]
    )

    doc.add_page_break()

    # ── Section 11 ─────────────────────────────────────────────────────────────
    heading1(doc, '11. Formatos de Archivo TXT (DGII)')
    normal(doc, (
        'Los archivos TXT generados siguen el formato oficial de la DGII. '
        'Se crean en /tmp/ como archivos temporales y se codifican en base64 '
        'para almacenarse en los campos binarios del modelo.'
    ))
    heading2(doc, '11.1 Estructura General')
    add_table(doc,
        ['Elemento', 'Descripción'],
        [
            ['Encabezado (línea 1)', '<tipo>|<RNC_empresa>|<YYYYMM>|<cantidad_registros>'],
            ['Líneas de datos', 'Campos separados por pipe (|), cantidades justificadas a la izquierda en 12 caracteres.'],
            ['Codificación', 'UTF-8'],
            ['Nombre de archivo', 'DGII_6XX_<RNC_empresa>_<YYYYMM>.txt'],
        ],
        col_widths=[2.0, 4.5]
    )
    heading2(doc, '11.2 Campos por Reporte')
    add_table(doc,
        ['Reporte', 'Campos por línea (en orden)'],
        [
            ['606', 'RNC/Cédula | Tipo ID | Tipo Gasto | NCF | NCF Modificado | Fecha NCF | Fecha Pago | Servicios | Bienes | Total | ITBIS Facturado | ITBIS Retenido | ITBIS Proporcionalidad | ITBIS Costo | ITBIS Adelanto | Tipo Retención ISR | Retención ISR | ISC | Otros impuestos | Propina Legal | Forma de Pago'],
            ['607', 'RNC/Cédula | Tipo ID | NCF | NCF Modificado | Tipo Ingreso | Fecha NCF | Fecha Retención | Total Facturado | ITBIS Facturado | ITBIS 3eros | ISR 3eros | ISC | Otros impuestos | Propina Legal | Efectivo | Cheque/Transf. | Tarjeta | Crédito | Bono | Permuta | Otros'],
            ['608', 'NCF Anulado | Fecha | Tipo Anulación'],
            ['609', 'Nombre Legal | Tipo ID | No. ID | Código País | Tipo Servicio | Detalle Servicio | Parte Relacionada | No. Documento | Fecha Doc. | Monto Facturado | Renta Presunta | ISR Retenido'],
        ],
        col_widths=[0.8, 5.7]
    )

    doc.add_page_break()

    # ── Section 12 ─────────────────────────────────────────────────────────────
    heading1(doc, '12. Pruebas Unitarias')
    heading2(doc, '12.1 Ejecución de la Suite de Pruebas')
    normal(doc, 'Para ejecutar las pruebas del módulo desde la línea de comandos:')
    code_block(doc, '$ python odoo-bin -d <base_datos> --test-enable --stop-after-init -i dgii_reports')
    normal(doc, 'Para ejecutar solo las pruebas de un archivo específico:')
    code_block(doc, '$ python odoo-bin -d <base_datos> --test-enable --stop-after-init --test-file=dgii_reports/tests/test_dgii_reports.py')

    heading2(doc, '12.2 Estructura de Pruebas')
    normal(doc, 'Los archivos de prueba se encuentran en el directorio tests/ del módulo:')
    add_table(doc,
        ['Archivo', 'Descripción'],
        [
            ['tests/__init__.py', 'Registro de módulos de prueba.'],
            ['tests/test_dgii_reports.py', 'Pruebas del modelo principal dgii.reports y generación de reportes.'],
        ],
        col_widths=[2.5, 4.0]
    )

    doc.add_page_break()

    # ── Section 13 ─────────────────────────────────────────────────────────────
    heading1(doc, '13. Configuración del Entorno de Desarrollo')
    heading2(doc, '13.1 Requisitos Previos')
    add_table(doc,
        ['Requisito', 'Versión mínima'],
        [
            ['Odoo', '19.0'],
            ['Python', '3.10+'],
            ['PostgreSQL', '14+'],
            ['pycountry (pip)', '22.x'],
            ['Módulos Odoo', 'l10n_do, l10n_do_accounting, accountant'],
        ],
        col_widths=[2.5, 4.0]
    )

    heading2(doc, '13.2 Instalación')
    numbered(doc, 'Clonar o copiar el módulo en el directorio de addons de Odoo.')
    numbered(doc, 'Instalar la dependencia Python: pip install pycountry')
    numbered(doc, 'Actualizar la lista de módulos en Odoo: Ajustes → Técnico → Actualizar lista de módulos.')
    numbered(doc, 'Buscar "Declaraciones DGII" en la lista de módulos e instalar.')
    numbered(doc, 'Verificar que el post_init_hook se haya ejecutado revisando los impuestos (campo Tipo DGII debe estar poblado).')

    heading2(doc, '13.3 Actualización del Módulo')
    code_block(doc, '$ python odoo-bin -d <base_datos> -u dgii_reports')

    doc.add_page_break()

    # ── Section 14 ─────────────────────────────────────────────────────────────
    heading1(doc, '14. Variables de Contexto y Depuración')
    heading2(doc, '14.1 Variables de Contexto Relevantes')
    add_table(doc,
        ['Variable de contexto', 'Uso en el módulo'],
        [
            ['allowed_company_ids', 'Utilizada en _get_companies_ids() para consolidar sucursales. Se sobreescribe con las IDs de compañías del grupo.'],
            ['company_id', 'Compañía activa. Usada como filtro base en consultas de facturas.'],
        ],
        col_widths=[2.2, 4.3]
    )

    heading2(doc, '14.2 Depuración')
    normal(doc, 'Para forzar el recómputo de campos DGII en facturas existentes:')
    code_block(doc, (
        '# Desde la consola de Odoo o shell:\n'
        'invoices = env["account.move"].search([("move_type", "in", ["in_invoice", "out_invoice"])])\n'
        'invoices.norma_recompute()\n'
    ))
    normal(doc, 'Para verificar el log del post_init_hook:')
    code_block(doc, '$ grep "dgii_reports" /var/log/odoo/odoo.log | grep -i "hook\\|tax_type\\|payment_form"')

    doc.save('doc/doc_tecnico_dgii_reports.docx')
    print('Technical doc saved.')


# ═══════════════════════════════════════════════════════════════════════════════
#  USER MANUAL
# ═══════════════════════════════════════════════════════════════════════════════

def build_user_manual():
    doc = Document('doc/doc_usuario_l10n_do_accounting.docx')
    # Start fresh keeping styles — preserve sectPr
    body = doc.element.body
    sectPr = body.find(qn('w:sectPr'))
    for element in list(body):
        body.remove(element)
    if sectPr is not None:
        body.append(sectPr)

    # ── Cover ─────────────────────────────────────────────────────────────────
    add_cover(
        doc,
        'Manual de Usuario',
        'Módulo de Declaraciones DGII\nRepública Dominicana — Odoo 19.0',
        'GUÍA DE INSTALACIÓN Y USO PARA USUARIOS FINALES',
        'Documento N.° LFL-USR-002',
        'GUÍA DE INSTALACIÓN Y USO PARA USUARIOS FINALES',
    )

    # ── TOC ───────────────────────────────────────────────────────────────────
    heading1(doc, 'Tabla de Contenido')
    toc = [
        '1.  Introducción',
        '2.  Instalación del Módulo',
        '    2.1  Requisitos Previos',
        '    2.2  Procedimiento de Instalación',
        '3.  Configuración Inicial',
        '    3.1  Configuración de Impuestos (Tipo DGII)',
        '    3.2  Configuración de Diarios (Forma de Pago)',
        '    3.3  Configuración de Cuentas (Anexo A e IT-1)',
        '    3.4  Consolidación Multi-Sucursal',
        '    3.5  Configuración de Proveedores Exterior',
        '4.  Creación de una Declaración DGII',
        '5.  El Reporte 606 — Compras y Gastos',
        '6.  El Reporte 607 — Ventas',
        '7.  El Reporte 608 — Comprobantes Anulados',
        '8.  El Reporte 609 — Pagos al Exterior',
        '9.  Anexo A e IT-1',
        '10. Resumen de Comprobantes Consumidor',
        '11. Descarga de Archivos TXT',
        '12. Marcar Declaración como Enviada',
        '13. Estados de las Facturas (Fiscal Status)',
        '14. Campos DGII en Facturas',
        '15. Preguntas Frecuentes',
        '16. Ambiente de Demostración',
    ]
    for e in toc:
        normal(doc, e)
    doc.add_page_break()

    # ── Section 1 ─────────────────────────────────────────────────────────────
    heading1(doc, '1. Introducción')
    normal(doc, (
        'El presente manual constituye la guía oficial de instalación y uso del módulo '
        'de Declaraciones DGII para Odoo 19.0, desarrollado para la República Dominicana. '
        'Este módulo automatiza la preparación de los reportes fiscales obligatorios ante la '
        'Dirección General de Impuestos Internos (DGII), eliminando la necesidad de preparar '
        'manualmente los archivos TXT requeridos.'
    ))
    normal(doc, 'Alcance del módulo:')
    for item in [
        'Generación automática del reporte 606 (Compras y Gastos).',
        'Generación automática del reporte 607 (Ventas).',
        'Generación automática del reporte 608 (Comprobantes Anulados).',
        'Generación automática del reporte 609 (Pagos al Exterior — facturas B17).',
        'Cálculo del Anexo A (IT-1) con 56 casillas agrupadas.',
        'Cálculo del formulario IT-1 completo con 68 casillas.',
        'Resumen de comprobantes consumidor (B02) para la declaración mensual.',
        'Marcado automático de facturas como "reportadas" al enviar la declaración.',
        'Descarga de todos los archivos TXT desde una sola pantalla.',
        'Consolidación de sucursales en un único reporte mensual.',
    ]:
        bullet(doc, item)
    normal(doc, 'Convenciones utilizadas en este documento:')
    bullet(doc, 'Las rutas de navegación se presentan en el formato: Menú → Sub-menú → Opción.')
    bullet(doc, 'Los nombres de campos y botones aparecen entre comillas ("Campo").')
    bullet(doc, 'Las notas importantes aparecen en cursiva precedidas por la palabra NOTA.')
    bullet(doc, 'Los ejemplos de valores aparecen en formato de código monoespaciado.')

    doc.add_page_break()

    # ── Section 2 ─────────────────────────────────────────────────────────────
    heading1(doc, '2. Instalación del Módulo')
    heading2(doc, '2.1 Requisitos Previos')
    normal(doc, 'Antes de instalar, verificar que se cumplan los siguientes requisitos:')
    add_table(doc,
        ['Requisito', 'Descripción'],
        [
            ['Odoo 19.0', 'Versión compatible. No válido en versiones anteriores.'],
            ['l10n_do instalado', 'Módulo base de localización dominicana.'],
            ['l10n_do_accounting instalado', 'Módulo de comprobantes fiscales (NCF Batches).'],
            ['pycountry (Python)', 'Librería para códigos ISO de países. Instalar con: pip install pycountry'],
            ['RNC de la compañía configurado', 'Ajustes → Compañía → campo NIF/RNC debe estar lleno.'],
        ],
        col_widths=[2.2, 4.3]
    )

    heading2(doc, '2.2 Procedimiento de Instalación')
    numbered(doc, 'Iniciar sesión en Odoo con una cuenta de usuario con rol Administrador.')
    numbered(doc, 'Desde el menú principal, hacer clic en la opción "Aplicaciones".')
    numbered(doc, 'En el campo de búsqueda, escribir el término: dgii_reports')
    numbered(doc, 'En el resultado, localizar el módulo "Declaraciones DGII".')
    numbered(doc, 'Hacer clic en el botón "Instalar".')
    numbered(doc, 'El sistema instalará automáticamente las dependencias requeridas.')
    numbered(doc, 'Al finalizar, verificar en los impuestos contables que el campo "Tipo DGII" esté correctamente asignado.')
    add_image_placeholder(doc, 'Búsqueda del módulo Declaraciones DGII en la tienda de aplicaciones')
    add_image_placeholder(doc, 'Módulo instalado — estado confirmado')

    doc.add_page_break()

    # ── Section 3 ─────────────────────────────────────────────────────────────
    heading1(doc, '3. Configuración Inicial')
    normal(doc, (
        'Antes de generar las declaraciones, es necesario verificar y completar la configuración '
        'de impuestos, diarios y cuentas contables.'
    ))

    heading2(doc, '3.1 Configuración de Impuestos (Tipo DGII)')
    normal(doc, (
        'Cada impuesto contable debe tener asignado un "Tipo DGII" para que el módulo '
        'pueda clasificar los montos correctamente en los reportes. El post_init_hook del '
        'módulo asigna estos tipos automáticamente al instalar, pero es posible verificarlos '
        'o ajustarlos manualmente.'
    ))
    normal(doc, 'Para verificar o modificar el Tipo DGII de un impuesto:')
    numbered(doc, 'Ir a Contabilidad → Configuración → Impuestos.')
    numbered(doc, 'Abrir el impuesto que se desea revisar.')
    numbered(doc, 'En la pestaña "Opciones avanzadas", verificar el campo "Tipo DGII".')
    numbered(doc, 'Si es un impuesto de retención ISR, también verificar el campo "Tipo Retención ISR".')
    add_image_placeholder(doc, 'Campo Tipo DGII y Tipo Retención ISR en el formulario de impuesto')
    add_table(doc,
        ['Tipo DGII', 'Descripción', 'Ejemplos de impuestos'],
        [
            ['itbis', 'ITBIS (18%, 16%, 9%, 8%, 0%)', 'ITBIS 18% ventas, ITBIS 18% compras'],
            ['ritbis', 'Retención de ITBIS', 'Retención 100% ITBIS personas físicas, Retención 30%'],
            ['isr', 'Retención de ISR (Impuesto Sobre la Renta)', 'Ret. 10% honorarios, Ret. 5% proveedores estado'],
            ['isc', 'Impuesto Selectivo al Consumo', 'ISC telecomunicaciones 10%'],
            ['tip', 'Propina legal (10%)', 'Propina legal ventas, Propina legal compras'],
            ['other', 'Otros impuestos', 'Gravamen 0.15% banco, Impuesto 2% telecomunicaciones'],
            ['none', 'Sin clasificación DGII', 'Impuestos internos no reportables'],
        ],
        col_widths=[1.2, 2.0, 3.3]
    )

    heading2(doc, '3.2 Configuración de Diarios (Forma de Pago)')
    normal(doc, (
        'Cada diario contable utilizado para registrar pagos debe tener configurada '
        'la "Forma de Pago" para que el módulo pueda determinar cómo clasificar los '
        'pagos en los reportes 606 y 607.'
    ))
    normal(doc, 'Para configurar la forma de pago de un diario:')
    numbered(doc, 'Ir a Contabilidad → Configuración → Diarios.')
    numbered(doc, 'Abrir el diario correspondiente (ej. "Efectivo", "Banco Principal").')
    numbered(doc, 'En la pestaña "Configuración Avanzada", localizar el campo "Forma de Pago".')
    numbered(doc, 'Seleccionar la forma de pago correspondiente.')
    numbered(doc, 'Guardar los cambios.')
    add_image_placeholder(doc, 'Campo Forma de Pago en el formulario de diario contable')
    add_table(doc,
        ['Opción', 'Descripción'],
        [
            ['Efectivo', 'Pagos en efectivo (código 01 en los reportes).'],
            ['Cheque/Transferencia/Depósito', 'Pagos bancarios, transferencias electrónicas (código 02).'],
            ['Tarjeta Crédito/Débito', 'Pagos con tarjeta de crédito o débito (código 03).'],
            ['Crédito', 'Ventas o compras a crédito (código 04).'],
            ['Permuta', 'Intercambios de bienes o servicios (código 05).'],
            ['Bono/Certificado', 'Pagos con bonos o certificados de regalo (código 06 — en ventas).'],
            ['Otros', 'Cualquier otra forma de pago no categorizada (código 07).'],
        ],
        col_widths=[2.2, 4.3]
    )

    heading2(doc, '3.3 Configuración de Cuentas (Anexo A e IT-1)')
    normal(doc, (
        'Para que el Anexo A e IT-1 se calculen automáticamente, es necesario asignar '
        'las cuentas contables a las casillas correspondientes del formulario. Esta es una '
        'configuración que generalmente se realiza una sola vez.'
    ))
    normal(doc, 'Para asignar una cuenta a una casilla:')
    numbered(doc, 'Ir a Contabilidad → Configuración → Plan de Cuentas.')
    numbered(doc, 'Abrir la cuenta que se desea configurar.')
    numbered(doc, 'En el campo "Casilla Anexo A (IT1)", seleccionar la casilla correspondiente para el Anexo A.')
    numbered(doc, 'En el campo "Casilla IT-1", seleccionar la casilla correspondiente para el IT-1.')
    numbered(doc, 'Guardar los cambios.')
    add_image_placeholder(doc, 'Campos Casilla Anexo A e IT-1 en el formulario de cuenta contable')
    normal(doc, (
        'NOTA: Los saldos del mes de todas las cuentas asignadas a una misma casilla '
        'se suman automáticamente al calcular el Anexo A e IT-1.'
    ))

    heading2(doc, '3.4 Consolidación Multi-Sucursal')
    normal(doc, (
        'Si la empresa tiene múltiples compañías en Odoo (sucursales), el módulo puede '
        'consolidar los datos de todas ellas en un único reporte mensual.'
    ))
    normal(doc, 'Para activar la consolidación:')
    numbered(doc, 'Ir a Ajustes → Compañías → Seleccionar la compañía principal (raíz).')
    numbered(doc, 'En la pestaña "DGII" (o "Declaraciones"), activar la opción "Consolidar Sucursales".')
    numbered(doc, 'Guardar los cambios.')
    add_image_placeholder(doc, 'Opción Consolidar Sucursales en configuración de compañía principal')
    normal(doc, (
        'NOTA: Esta opción solo está disponible en la compañía raíz (padre). '
        'Las compañías sucursal mostrarán un mensaje informativo en lugar del toggle.'
    ))

    heading2(doc, '3.5 Configuración de Proveedores del Exterior (B17)')
    normal(doc, (
        'Para que las facturas de proveedores del exterior se incluyan correctamente en '
        'el reporte 609, es necesario configurar el tipo de servicio en cada factura.'
    ))
    normal(doc, 'Para indicar que un proveedor es parte relacionada:')
    numbered(doc, 'Ir al formulario del contacto del proveedor extranjero.')
    numbered(doc, 'En la pestaña "Ventas y Compras", localizar el campo "Parte Relacionada".')
    numbered(doc, 'Seleccionar "Relacionado" si el proveedor es parte vinculada a la empresa.')
    add_image_placeholder(doc, 'Campo Parte Relacionada en formulario de contacto proveedor exterior')

    doc.add_page_break()

    # ── Section 4 ─────────────────────────────────────────────────────────────
    heading1(doc, '4. Creación de una Declaración DGII')
    normal(doc, (
        'Una declaración DGII agrupa los cuatro reportes (606, 607, 608, 609) para un '
        'período mensual específico. Se crea uno por mes y por compañía.'
    ))
    normal(doc, 'Para crear una nueva declaración:')
    numbered(doc, 'Ir a Contabilidad → Declaraciones DGII (o desde el menú DGII si está configurado).')
    numbered(doc, 'Hacer clic en "Nuevo".')
    numbered(doc, 'En el campo "Período", ingresar el mes y año en formato MM/YYYY. Ejemplo: 03/2025.')
    numbered(doc, 'El sistema calculará automáticamente la fecha de inicio y fin del período.')
    numbered(doc, 'Si aplica, ingresar el "Saldo Anterior" (balance del período anterior).')
    numbered(doc, 'Hacer clic en "Guardar".')
    numbered(doc, 'Hacer clic en el botón "Generar Declaraciones" para procesar los datos.')
    add_image_placeholder(doc, 'Formulario de nueva declaración DGII con campo de período')
    add_image_placeholder(doc, 'Dashboard de declaración generada mostrando resumen de los 4 reportes')
    normal(doc, (
        'NOTA: Si ya existe una declaración para el mismo período y compañía, el sistema '
        'mostrará un error de validación. Solo puede existir una declaración por período.'
    ))
    normal(doc, (
        'NOTA: Si la declaración ya fue generada anteriormente y se hace clic en "Generar '
        'Declaraciones" de nuevo, el sistema mostrará un diálogo de confirmación antes de '
        'borrar y recalcular. Confirmar solo si se desea regenerar completamente.'
    ))
    add_image_placeholder(doc, 'Diálogo de confirmación de regeneración de declaración')

    doc.add_page_break()

    # ── Section 5 ─────────────────────────────────────────────────────────────
    heading1(doc, '5. El Reporte 606 — Compras y Gastos')
    normal(doc, (
        'El reporte 606 contiene el registro de todas las compras y gastos del período, '
        'incluyendo facturas de proveedores con NCF. Es el equivalente al "Libro de Compras".'
    ))
    normal(doc, 'Facturas incluidas en el 606:')
    bullet(doc, 'Facturas de proveedor (in_invoice) confirmadas en el período.')
    bullet(doc, 'Notas de crédito de proveedor (in_refund) confirmadas en el período.')
    bullet(doc, 'Facturas de períodos anteriores que fueron pagadas en el período actual.')
    bullet(doc, 'NO se incluyen facturas con estado fiscal "No enviada" (blocked).')

    normal(doc, 'Para ver el detalle del reporte 606:')
    numbered(doc, 'Abrir la declaración DGII del período deseado.')
    numbered(doc, 'En el panel de resumen, hacer clic sobre el recuadro "606".')
    numbered(doc, 'Se abrirá la vista de lista con todas las líneas del reporte.')
    numbered(doc, 'Hacer clic en cualquier línea para ver los detalles o abrir la factura original.')
    add_image_placeholder(doc, 'Vista de lista del reporte 606 con todas las líneas de compras')
    add_image_placeholder(doc, 'Detalle de una línea 606 mostrando todos los campos DGII')
    normal(doc, 'Campos principales visibles en el listado:')
    add_table(doc,
        ['Campo', 'Descripción'],
        [
            ['RNC/Cédula', 'Identificación fiscal del proveedor.'],
            ['Tipo de Gasto', 'Categoría del gasto (01-Bienes para producción, 02-Bienes para reventa, etc.).'],
            ['NCF', 'Número de comprobante fiscal del proveedor.'],
            ['Fecha NCF', 'Fecha de emisión de la factura.'],
            ['Fecha Pago', 'Fecha del último pago registrado.'],
            ['Servicios', 'Monto de servicios sin ITBIS.'],
            ['Bienes', 'Monto de bienes sin ITBIS.'],
            ['ITBIS Facturado', 'ITBIS declarado en la factura.'],
            ['ITBIS Retenido', 'RITBIS retenido al proveedor.'],
            ['Tipo Ret. ISR', 'Categoría de retención ISR (01-08).'],
            ['Retención ISR', 'Monto de ISR retenido.'],
            ['Forma de Pago', 'Método de pago utilizado.'],
        ],
        col_widths=[1.8, 4.7]
    )

    doc.add_page_break()

    # ── Section 6 ─────────────────────────────────────────────────────────────
    heading1(doc, '6. El Reporte 607 — Ventas')
    normal(doc, (
        'El reporte 607 contiene el registro de todas las ventas del período con NCF. '
        'Es el equivalente al "Libro de Ventas".'
    ))
    normal(doc, 'Facturas incluidas en el 607:')
    bullet(doc, 'Facturas de cliente (out_invoice) confirmadas en el período.')
    bullet(doc, 'Notas de crédito de cliente (out_refund) confirmadas en el período.')
    bullet(doc, 'Facturas de clientes de períodos anteriores pagadas en el período actual.')
    bullet(doc, 'Comprobantes consumidor (B02) con monto ≥ RD$250,000 se incluyen en el TXT.')
    bullet(doc, 'Comprobantes consumidor (B02) con monto < RD$250,000 se incluyen en el resumen pero NO en el TXT.')

    normal(doc, 'Para ver el detalle del reporte 607:')
    numbered(doc, 'Abrir la declaración del período.')
    numbered(doc, 'Hacer clic sobre el recuadro "607" en el panel de resumen.')
    add_image_placeholder(doc, 'Vista de lista del reporte 607 con todas las líneas de ventas')
    normal(doc, 'Campos principales en el listado 607:')
    add_table(doc,
        ['Campo', 'Descripción'],
        [
            ['RNC/Cédula', 'Identificación fiscal del cliente.'],
            ['NCF', 'Número de comprobante fiscal de la venta.'],
            ['Tipo de Ingreso', 'Categoría del ingreso (01-Operacional, 02-Financiero, etc.).'],
            ['Fecha NCF', 'Fecha de emisión de la factura.'],
            ['Monto Facturado', 'Total de la factura.'],
            ['ITBIS Facturado', 'ITBIS declarado.'],
            ['ITBIS Retenido 3eros', 'RITBIS retenido por terceros.'],
            ['ISR Retenido 3eros', 'ISR retenido por terceros.'],
            ['Efectivo / Cheque / Tarjeta…', 'Desglose del pago por forma de pago.'],
        ],
        col_widths=[2.0, 4.5]
    )

    doc.add_page_break()

    # ── Section 7 ─────────────────────────────────────────────────────────────
    heading1(doc, '7. El Reporte 608 — Comprobantes Anulados')
    normal(doc, (
        'El reporte 608 contiene todos los comprobantes fiscales que fueron anulados '
        'durante el período. Incluye facturas de venta, compra y notas de crédito de venta '
        'que se encuentren en estado "Cancelado".'
    ))
    normal(doc, 'Para ver el detalle del reporte 608:')
    numbered(doc, 'Abrir la declaración del período.')
    numbered(doc, 'Hacer clic sobre el recuadro "608" en el panel de resumen.')
    add_image_placeholder(doc, 'Vista de lista del reporte 608 — comprobantes anulados')
    add_table(doc,
        ['Campo', 'Descripción'],
        [
            ['NCF Anulado', 'Número de comprobante que fue anulado.'],
            ['Fecha', 'Fecha original de la factura anulada.'],
            ['Tipo de Anulación', 'Razón de la anulación según tipos DGII.'],
        ],
        col_widths=[2.0, 4.5]
    )

    doc.add_page_break()

    # ── Section 8 ─────────────────────────────────────────────────────────────
    heading1(doc, '8. El Reporte 609 — Pagos al Exterior')
    normal(doc, (
        'El reporte 609 incluye todos los pagos realizados a proveedores del exterior '
        'durante el período. Se identifican por el prefijo de NCF B17 en las facturas '
        'de proveedor.'
    ))
    normal(doc, 'Requisitos para que una factura aparezca en 609:')
    bullet(doc, 'Debe ser una factura de proveedor (in_invoice) confirmada.')
    bullet(doc, 'El NCF de la factura debe comenzar con "B17".')
    bullet(doc, 'La fecha de la factura debe caer dentro del período declarado.')
    normal(doc, 'Configuración del tipo de servicio en la factura:')
    numbered(doc, 'Abrir la factura del proveedor exterior.')
    numbered(doc, 'En la pestaña "DGII / Información Fiscal", localizar el campo "Tipo de Servicio".')
    numbered(doc, 'Seleccionar la categoría principal del servicio adquirido (01-08).')
    numbered(doc, 'En el campo "Detalle del Servicio", seleccionar la subcategoría específica.')
    add_image_placeholder(doc, 'Campos Tipo de Servicio y Detalle en factura de proveedor exterior (B17)')
    normal(doc, 'Para ver el detalle del reporte 609:')
    numbered(doc, 'Abrir la declaración del período.')
    numbered(doc, 'Hacer clic sobre el recuadro "609" en el panel de resumen.')
    add_image_placeholder(doc, 'Vista de lista del reporte 609 — pagos al exterior')
    add_table(doc,
        ['Campo', 'Descripción'],
        [
            ['Nombre Legal', 'Razón social del proveedor extranjero.'],
            ['Tipo ID / No. ID', 'Tipo y número de identificación fiscal del proveedor.'],
            ['Código País', 'Código numérico ISO 3166 del país del proveedor.'],
            ['Tipo Servicio', 'Categoría del servicio adquirido (01-08).'],
            ['Detalle Servicio', 'Subcategoría detallada.'],
            ['Parte Relacionada', '0=No relacionado, 1=Parte vinculada.'],
            ['Monto Facturado', 'Total de la factura en moneda local.'],
            ['ISR Retenido', 'ISR retenido al proveedor exterior.'],
        ],
        col_widths=[2.0, 4.5]
    )

    doc.add_page_break()

    # ── Section 9 ─────────────────────────────────────────────────────────────
    heading1(doc, '9. Anexo A e IT-1')
    normal(doc, (
        'El Anexo A e IT-1 son formularios complementarios que agrupan los datos de '
        'ventas y compras en casillas predefinidas por la DGII. Se calculan automáticamente '
        'a partir de los saldos de las cuentas contables configuradas.'
    ))
    heading2(doc, '9.1 Anexo A (Secciones II-VI)')
    normal(doc, 'El Anexo A contiene 56 casillas agrupadas en 6 secciones:')
    add_table(doc,
        ['Sección', 'Descripción'],
        [
            ['II', 'Operaciones por tipo de NCF (Crédito Fiscal, Consumidor, Débito, Crédito, Único, Especial, Gubernamental, Exportación).'],
            ['III', 'Operaciones por forma de venta (Efectivo, Cheque, Tarjeta, Crédito, Bonos, Permuta, Otros).'],
            ['IV', 'Operaciones por tipo de ingreso (Operacionales, Financieros, Extraordinarios, Arrendamientos, Venta activos, Otros).'],
            ['V', 'Pagos y retenciones computables (Normas 08-04, 02-05, paquetes alojamiento, entes del estado, ITBIS percibido).'],
            ['VI', 'Operaciones de construcción y operaciones a comisión.'],
        ],
        col_widths=[0.8, 5.7]
    )
    heading2(doc, '9.2 IT-1 (Formulario de Declaración ITBIS)')
    normal(doc, 'El IT-1 contiene 68 casillas en múltiples secciones, incluyendo:')
    bullet(doc, 'Sección II: Ingresos por tipo de operación (exentos y gravados a 18%, 16%, 9%, 8%).')
    bullet(doc, 'Sección III: Liquidación (ITBIS cobrado, deducible, a pagar, pagos computables).')
    bullet(doc, 'Sección IV: Recargos, intereses y sanciones.')
    bullet(doc, 'Sección V-A: ITBIS retenido/percibido (personas físicas, ONGs, empresas, RST).')
    bullet(doc, 'Sección V-B: Recargos por retención.')
    bullet(doc, 'Sección V-C: Total a pagar.')
    add_image_placeholder(doc, 'Pestaña Anexo A en formulario de declaración DGII')
    add_image_placeholder(doc, 'Pestaña IT-1 en formulario de declaración DGII')
    normal(doc, (
        'NOTA: Los valores del Anexo A e IT-1 son editables. El usuario puede '
        'ajustar manualmente cualquier casilla si el cálculo automático no refleja '
        'correctamente la realidad de la declaración.'
    ))

    doc.add_page_break()

    # ── Section 10 ─────────────────────────────────────────────────────────────
    heading1(doc, '10. Resumen de Comprobantes Consumidor')
    normal(doc, (
        'La pestaña "Resumen Consumidor" de la declaración muestra el consolidado de '
        'todos los comprobantes B02 (Consumidor Final) emitidos durante el período. '
        'Esta información se incluye en la declaración sin que los comprobantes '
        'menores a RD$250,000 aparezcan individualmente en el TXT 607.'
    ))
    add_image_placeholder(doc, 'Pestaña Resumen Consumidor en formulario de declaración DGII')
    add_table(doc,
        ['Campo', 'Descripción'],
        [
            ['Cantidad NCF', 'Total de comprobantes consumidor emitidos en el período.'],
            ['Monto Total', 'Suma de todos los montos de facturas B02.'],
            ['ITBIS Total', 'ITBIS total de facturas B02.'],
            ['ISC Total', 'ISC total de facturas B02.'],
            ['Otros Impuestos', 'Otros impuestos en facturas B02.'],
            ['Propina Legal', 'Propina legal total en facturas B02.'],
            ['Efectivo / Cheque / Tarjeta / Crédito / Bono / Permuta / Otros', 'Desglose del total B02 por forma de pago.'],
        ],
        col_widths=[2.5, 4.0]
    )

    doc.add_page_break()

    # ── Section 11 ─────────────────────────────────────────────────────────────
    heading1(doc, '11. Descarga de Archivos TXT')
    normal(doc, (
        'Una vez generada la declaración, los archivos TXT en el formato oficial de la DGII '
        'están disponibles para descarga en la pestaña "Archivos TXT".'
    ))
    normal(doc, 'Para descargar los archivos TXT:')
    numbered(doc, 'Abrir la declaración DGII generada.')
    numbered(doc, 'Hacer clic en la pestaña "Archivos TXT".')
    numbered(doc, 'Hacer clic en el botón de descarga junto al reporte deseado (606, 607, 608, 609).')
    numbered(doc, 'El archivo se descargará con el formato: DGII_6XX_<RNC_empresa>_<YYYYMM>.txt')
    numbered(doc, 'Ingresar al portal DGII (dgii.gov.do) y cargar el archivo correspondiente.')
    add_image_placeholder(doc, 'Pestaña Archivos TXT con botones de descarga para los 4 reportes')
    add_table(doc,
        ['Archivo', 'Formato de nombre', 'Contenido'],
        [
            ['606', 'DGII_606_<RNC>_<YYYYMM>.txt', 'Compras y gastos del período.'],
            ['607', 'DGII_607_<RNC>_<YYYYMM>.txt', 'Ventas del período.'],
            ['608', 'DGII_608_<RNC>_<YYYYMM>.txt', 'Comprobantes anulados del período.'],
            ['609', 'DGII_609_<RNC>_<YYYYMM>.txt', 'Pagos al exterior del período.'],
        ],
        col_widths=[0.8, 2.4, 3.3]
    )

    doc.add_page_break()

    # ── Section 12 ─────────────────────────────────────────────────────────────
    heading1(doc, '12. Marcar Declaración como Enviada')
    normal(doc, (
        'Después de subir los archivos TXT al portal de la DGII, es importante marcar '
        'la declaración como enviada en Odoo. Esta acción actualiza el estado fiscal '
        'de todas las facturas incluidas en el reporte.'
    ))
    normal(doc, 'Para marcar como enviada:')
    numbered(doc, 'Abrir la declaración DGII con estado "Generado".')
    numbered(doc, 'Hacer clic en el botón "Marcar como Enviado" en el encabezado del formulario.')
    numbered(doc, 'Confirmar la acción si se solicita.')
    numbered(doc, 'El estado de la declaración cambiará a "Enviado".')
    numbered(doc, 'Todas las facturas incluidas en la declaración tendrán su "Estado Fiscal" actualizado a "Reportada"..')
    add_image_placeholder(doc, 'Botón "Marcar como Enviado" y estado Enviado en barra de estado')
    normal(doc, (
        'NOTA: Una vez marcada como enviada, la declaración no puede regenerarse '
        'directamente. Si necesita corregir datos, contacte al administrador del sistema.'
    ))

    doc.add_page_break()

    # ── Section 13 ─────────────────────────────────────────────────────────────
    heading1(doc, '13. Estados de las Facturas (Fiscal Status)')
    normal(doc, (
        'El módulo agrega un campo "Estado Fiscal" a todas las facturas confirmadas '
        'que tienen NCF. Este campo controla si la factura ha sido declarada ante la DGII.'
    ))
    add_table(doc,
        ['Estado', 'Descripción', 'Acción requerida'],
        [
            ['Parcial (normal)', 'Factura pendiente de ser declarada. Estado inicial al confirmar la factura.', 'Generar y enviar la declaración del período correspondiente.'],
            ['Reportada (done)', 'Factura incluida en una declaración marcada como enviada.', 'Ninguna. La factura ya fue declarada.'],
            ['No enviada (blocked)', 'Factura excluida manualmente de las declaraciones.', 'Revisar con el equipo contable si procede incluirla.'],
        ],
        col_widths=[1.5, 2.8, 2.2]
    )
    add_image_placeholder(doc, 'Campo Estado Fiscal en listado de facturas')
    normal(doc, (
        'NOTA: El campo "Estado Fiscal" es visible en la vista de lista de facturas '
        'y permite filtrar rápidamente las facturas pendientes de declarar.'
    ))

    doc.add_page_break()

    # ── Section 14 ─────────────────────────────────────────────────────────────
    heading1(doc, '14. Campos DGII en Facturas')
    normal(doc, (
        'El módulo agrega una pestaña "DGII / Información Fiscal" en el formulario de '
        'facturas que muestra todos los campos calculados automáticamente para los reportes.'
    ))
    add_image_placeholder(doc, 'Pestaña DGII en formulario de factura de proveedor')
    add_image_placeholder(doc, 'Pestaña DGII en formulario de factura de cliente')
    add_table(doc,
        ['Campo', 'Tipo de factura', 'Descripción'],
        [
            ['ITBIS Facturado', 'Compras y Ventas', 'Monto de ITBIS declarado en la factura.'],
            ['ITBIS Retenido', 'Compras', 'RITBIS retenido al proveedor.'],
            ['ISR Retenido', 'Compras', 'ISR retenido al proveedor.'],
            ['ISC', 'Compras y Ventas', 'Impuesto Selectivo al Consumo.'],
            ['Propina Legal', 'Compras y Ventas', 'Propina legal del 10%.'],
            ['Monto Servicios', 'Compras', 'Subtotal de líneas de servicios sin ITBIS.'],
            ['Monto Bienes', 'Compras', 'Subtotal de líneas de bienes sin ITBIS.'],
            ['Forma de Pago', 'Compras', 'Método de pago registrado para el 606.'],
            ['Fecha de Pago', 'Compras', 'Fecha del último pago registrado.'],
            ['Tipo Retención ISR', 'Compras', 'Categoría de retención ISR (01-08).'],
            ['Tipo de Servicio', 'Compras (B17)', 'Categoría de servicio para el 609.'],
            ['Detalle de Servicio', 'Compras (B17)', 'Subcategoría para el 609.'],
            ['Estado Fiscal', 'Compras y Ventas', 'normal / done / blocked.'],
        ],
        col_widths=[2.0, 1.5, 3.0]
    )

    doc.add_page_break()

    # ── Section 15 ─────────────────────────────────────────────────────────────
    heading1(doc, '15. Preguntas Frecuentes')
    add_table(doc,
        ['Pregunta', 'Respuesta'],
        [
            ['¿Puedo generar el reporte más de una vez en el mismo período?',
             'Sí. Al hacer clic en "Generar Declaraciones" en un reporte ya generado, el sistema mostrará un diálogo de confirmación. Al confirmar, se borran todas las líneas anteriores y se recalcula desde cero con las facturas actuales.'],
            ['¿Por qué no aparece una factura en el reporte?',
             'Verificar: (1) que la factura esté en estado "Publicado/Confirmado", (2) que la fecha de la factura esté dentro del período declarado, (3) que el estado fiscal no sea "blocked", (4) que el NCF esté correcto.'],
            ['¿Qué pasa con las facturas pagadas en un mes diferente al de su emisión?',
             'Se incluyen en la declaración del mes en que fueron pagadas, no del mes de emisión, siempre que su estado fiscal sea "normal".'],
            ['¿Por qué el archivo TXT 607 tiene menos registros que la lista en pantalla?',
             'Los comprobantes consumidor (B02) con monto menor a RD$250,000 se muestran en la lista pero no se incluyen en el TXT. Solo se incluyen en el resumen de consumidores.'],
            ['¿Cómo corrijo una factura después de generar la declaración?',
             'Corregir la factura en Odoo, luego regenerar la declaración haciendo clic en "Generar Declaraciones" y confirmando la regeneración.'],
            ['¿Qué significa "Saldo Anterior"?',
             'Es el balance a favor o en contra que se traslada del período anterior. Se ingresa manualmente consultando la declaración anterior.'],
            ['¿Cómo incluyo una factura de proveedor exterior en el 609?',
             'La factura debe tener un NCF que comience con "B17". Adicionalmente, debe tener configurado el "Tipo de Servicio" y "Detalle de Servicio" en la pestaña DGII.'],
            ['¿El Anexo A e IT-1 se calculan solos?',
             'Sí, siempre que las cuentas contables estén correctamente asignadas a las casillas del formulario. Ver sección 3.3 de este manual.'],
            ['¿Puedo editar los valores del Anexo A e IT-1?',
             'Sí. Los valores son editables directamente en la pestaña correspondiente de la declaración. Los cambios se guardan en la declaración sin afectar las cuentas contables.'],
            ['¿Qué pasa si olvido marcar la declaración como enviada?',
             'Las facturas permanecerán en estado "Parcial" (normal) y podrán aparecer en la siguiente declaración si coincide su fecha de pago. Se recomienda siempre marcar como enviada después de subir los archivos a la DGII.'],
        ],
        col_widths=[2.4, 4.1]
    )

    doc.add_page_break()

    # ── Section 16 ─────────────────────────────────────────────────────────────
    heading1(doc, '16. Ambiente de Demostración')
    normal(doc, (
        'El módulo dgii_reports incluye datos de demostración comprensivos que permiten '
        'evaluar y validar el funcionamiento de todos los reportes DGII sin necesidad de '
        'ingresar datos manualmente. Al instalar Odoo con datos demo activados, el sistema '
        'configura automáticamente una compañía completa lista para usar.'
    ))

    heading2(doc, '16.1 Cómo Activar los Datos Demo')
    normal(doc, 'Los datos demo se cargan automáticamente cuando se instala Odoo con la opción de datos de demostración:')
    add_table(doc,
        ['Método', 'Instrucción'],
        [
            ['Instalación nueva', 'Al crear la base de datos en el asistente de Odoo, activar la opción "Cargar datos de demostración".'],
            ['Línea de comandos', 'Ejecutar: odoo-bin -d nombre_db --without-demo=False -i dgii_reports'],
        ],
        col_widths=[1.8, 4.7]
    )
    normal(doc, (
        'NOTA: Los datos demo requieren que l10n_do_accounting también esté instalado con datos demo. '
        'La compañía demo se llama "DISTRIBUIDORA COMERCIAL DOMINICANA SRL".'
    ))

    heading2(doc, '16.2 Qué se Crea Automáticamente')
    normal(doc, 'Al activar los datos demo, se crea lo siguiente:')

    heading3(doc, 'Compañía y Configuración Base')
    add_table(doc,
        ['Elemento', 'Valor / Descripción'],
        [
            ['Compañía', 'DISTRIBUIDORA COMERCIAL DOMINICANA SRL — RNC: 131552111'],
            ['País', 'República Dominicana — Moneda DOP'],
            ['Plan de cuentas', 'Plan de cuentas dominicano completo (l10n_do)'],
            ['NCF Batches', 'Lotes configurados para todos los tipos de comprobante (B01, B02, B03, B04, B11, B13, B14, B15, B16)'],
            ['Diario de tarjeta', 'Diario "Tarjeta de Crédito/Débito" (código: CARD) para registrar pagos con tarjeta'],
            ['Tipos de impuesto DGII', 'Todos los impuestos configurados: ITBIS, RITBIS, ISR (con tipo de retención), ISC, propina legal'],
        ],
        col_widths=[2.0, 4.5]
    )

    heading3(doc, 'Socios (Clientes y Proveedores)')
    add_table(doc,
        ['Tipo', 'Socios creados'],
        [
            ['Contribuyentes (RNC 9 dígitos)', 'MARCOS ORGANIZADOR DE NEGOCIOS SRL, TECNOLOGIAS DIGITALES DO SRL, KATANA LABS SRL, SOLUCIONES EMPRESARIALES DEL CARIBE SRL'],
            ['No contribuyentes (cédula 11 dígitos)', 'JOSE LUIS LOPEZ GONZALEZ, KEVIN JIMENEZ LORENZO, MARIA FERNANDA RODRIGUEZ CASTILLO'],
            ['Proveedor informal', 'JUAN BAUTISTA MARTE PEREZ'],
            ['Zona Franca', 'ZONA FRANCA INDUSTRIAL DE LAS AMERICAS, ZONA FRANCA INDUSTRIAL DE SANTIAGO'],
            ['Gubernamental', 'MINISTERIO DE INDUSTRIA Y COMERCIO, MINISTERIO DE EDUCACIÓN, MINISTERIO DE SALUD PÚBLICA'],
            ['ONG / Sin fines de lucro', 'FOOD FOR THE HUNGRY Y DOM, CARITAS DOMINICANA INC'],
            ['Extranjeros', 'Azure Interior (USA), SERVICIOS TECNOLOGICOS GLOBALES S.A. (España), SOLUCIONES DIGITALES DE COLOMBIA S.A.S. (Colombia)'],
        ],
        col_widths=[2.2, 4.3]
    )

    heading3(doc, 'Facturas de Ventas — Reporte 607')
    normal(doc, 'Se crean facturas para cada tipo de NCF de venta y cada tipo de ingreso:')
    add_table(doc,
        ['Tipo NCF', 'Detalle', 'Forma de Pago'],
        [
            ['B01 (Fiscal)', 'Tipos de ingreso 01 a 06 — una factura por cada tipo', 'Efectivo, banco, crédito'],
            ['B01 con RITBIS 30%', 'Retención de ITBIS por parte del cliente — valida columna "ITBIS Retenido 3eros" en 607', 'Sin pago'],
            ['B15 Gov. + ISR 5%', 'Retención de ISR gubernamental — valida columna "ISR Retenido 3eros" en 607', 'Sin pago'],
            ['B01 con ISC', 'Impuesto Selectivo al Consumo (telecomunicaciones) — valida columna ISC en 607', 'Banco'],
            ['B02 Consumidor (<250k)', 'Monto pequeño — aparece solo en resumen consumidor, no en TXT', 'Efectivo'],
            ['B02 Consumidor (≥250k)', 'Monto grande — se incluye en TXT 607', 'Banco'],
            ['B02 con Propina Legal', 'Factura de restaurante con tax_tip_sale — valida columna Propina en 607', 'Tarjeta'],
            ['B14 Régimen Especial', 'Zona Franca — 0% ITBIS, dos facturas', 'Sin pago'],
            ['B16 Exportación', 'Cliente extranjero — 0% ITBIS', 'Banco'],
            ['B04 Nota de Crédito', 'Devolución referenciando B01 — valida campo NCF Modificado', '—'],
            ['B03 Nota de Débito', 'Ajuste referenciando B01 — valida campo NCF Modificado', '—'],
        ],
        col_widths=[1.8, 3.3, 1.4]
    )

    heading3(doc, 'Facturas de Compras — Reporte 606')
    normal(doc, 'Se crean facturas cubriendo los 11 tipos de gasto y todas las combinaciones de impuesto relevantes:')
    add_table(doc,
        ['Tipo de Gasto', 'Descripción', 'Columnas que valida'],
        [
            ['01 — Personal', 'Servicios de limpieza (product=servicio)', 'service_total_amount, payment_type 02 (banco)'],
            ['02 — Trabajo/Insumos', 'Insumos de trabajo (product=bien)', 'good_total_amount, payment_type 01 (efectivo)'],
            ['02 — Con RITBIS+ISR', 'Honorarios + materiales — RITBIS 100% + ISR 10%', 'withholded_itbis, income_withholding, isr_type 02'],
            ['03 — Arrendamiento', 'Alquiler de local comercial', 'expense_type 03, payment_type 02'],
            ['04 — Activo Fijo', 'Equipo de cómputo (product=bien)', 'good_total_amount, payment_type 03 (tarjeta)'],
            ['05 — Representación', 'Gastos de representación', 'payment_type 04 (crédito/sin pago)'],
            ['06 — Deducciones', 'Construcción — RITBIS 75%', 'withholded_itbis (75%), expense_type 06'],
            ['07 — Financiero', 'Intereses bancarios', 'expense_type 07'],
            ['08 — Extraordinario', 'Gasto por siniestro', 'expense_type 08'],
            ['09 — Costo de Ventas', 'Insumos en costo', 'expense_type 09'],
            ['10 — Adquisición Activos', 'Activo intangible (product=bien)', 'expense_type 10'],
            ['11 — Seguros', 'Prima de seguro corporativo', 'expense_type 11'],
            ['02 — ISC', 'Servicio de telecomunicaciones — tax_10_telco', 'selective_tax (ISC)'],
            ['Nota de Crédito', 'Devolución de compra B11', 'modified_invoice_number, credit_note'],
        ],
        col_widths=[1.7, 2.3, 2.5]
    )

    heading3(doc, 'Comprobantes Anulados — Reporte 608')
    normal(doc, (
        'Se crean y confirman 10 facturas de venta, luego se cancelan con los 10 tipos de anulación '
        'definidos por la DGII (01 al 10), cubriendo todos los registros posibles del reporte 608.'
    ))

    heading3(doc, 'Pagos al Exterior — Reporte 609')
    add_table(doc,
        ['Tipo Servicio', 'País', 'Parte Relacionada', 'ISR 27%', 'Columnas que valida'],
        [
            ['01 — Importación bienes', 'España', 'No (0)', 'No', 'service_type, country_code'],
            ['02 — Servicios prof.', 'España', 'No (0)', 'Sí', 'withholded_isr, isr_withholding_date'],
            ['02 — Servicios prof.', 'USA', 'No (0)', 'No', 'sin ISR → isr_date vacío'],
            ['03 — Servicios técnicos', 'USA', 'No (0)', 'Sí', 'withholded_isr, isr_withholding_date'],
            ['03 — Servicios técnicos', 'Colombia', 'Sí (1)', 'No', 'related_part=1'],
            ['04 — Gastos financieros', 'España', 'No (0)', 'No', 'service_type 04'],
            ['05 — Seguros', 'USA', 'No (0)', 'Sí', 'service_type 05, withholded_isr'],
            ['06 — Comisiones', 'España', 'No (0)', 'No', 'service_type 06'],
            ['07 — Alquileres', 'Colombia', 'No (0)', 'No', 'service_type 07'],
            ['08 — Regalías', 'USA', 'Sí (1)', 'Sí', 'service_type 08, related_part=1'],
        ],
        col_widths=[1.6, 0.9, 1.1, 0.8, 2.6]
    )

    heading2(doc, '16.3 Cómo Explorar la Data Demo')
    numbered(doc, 'Ir a Contabilidad → DGII → Declaraciones DGII.')
    numbered(doc, 'Seleccionar la declaración generada (período del mes anterior).')
    numbered(doc, 'Hacer clic en las pestañas 606, 607, 608 y 609 para ver las líneas.')
    numbered(doc, 'Hacer clic en "Ver 606", "Ver 607", etc. para abrir las vistas árbol con todos los detalles.')
    numbered(doc, 'Descargar los archivos TXT y verificar el contenido de cada columna.')
    normal(doc, (
        'NOTA: La declaración ya viene generada para el período del mes anterior. '
        'Si se necesita regenerar, hacer clic en el botón "Generar Declaraciones".'
    ))

    heading2(doc, '16.4 Uso para Pruebas de Desarrollo')
    normal(doc, (
        'Los datos demo están diseñados específicamente para validar el cálculo de cada columna '
        'de cada reporte. Si un campo aparece en 0 cuando se esperaba un valor, se debe verificar:'
    ))
    add_table(doc,
        ['Columna vacía', 'Qué verificar'],
        [
            ['invoiced_itbis / withholded_itbis', 'Que l10n_do_tax_type esté en "itbis" o "ritbis" en los impuestos de la compañía.'],
            ['income_withholding / isr_withholding_type', 'Que l10n_do_tax_type="isr" e isr_retention_type estén configurados en los impuestos ISR.'],
            ['selective_tax', 'Que tax_10_telco exista en el plan de cuentas demo y l10n_do_tax_type="isc".'],
            ['legal_tip', 'Que tax_tip_sale exista y l10n_do_tax_type="tip".'],
            ['service_total_amount / good_total_amount', 'Que las líneas de factura tengan un producto con type="service" o type="consu".'],
            ['cash / bank / card en 607', 'Que los diarios cash/bank tengan l10n_do_payment_form configurado, y que exista el diario CARD.'],
            ['withholded_isr en 609', 'Que ret_27_income_remittance tenga l10n_do_tax_type="isr" y que la factura B17 esté pagada.'],
        ],
        col_widths=[2.2, 4.3]
    )

    doc.save('doc/doc_usuario_dgii_reports.docx')
    print('User manual saved.')


if __name__ == '__main__':
    build_technical_doc()
    build_user_manual()
    print('Done.')
