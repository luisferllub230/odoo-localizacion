#!/usr/bin/env python3
"""
Script de automatización para capturar screenshots del flujo completo
de l10n_do_accounting en Odoo y agregarlos al documento Word.
"""

import os
import time
import xmlrpc.client
from pathlib import Path
from playwright.sync_api import sync_playwright, Page

# Configuración
URL = "http://localhost:8092"
DB = "test_v19e"
ADMIN_USER = "admin"
ADMIN_PASS = "admin"
SCREENSHOTS_DIR = Path("doc/screenshots")
SCREENSHOTS_DIR.mkdir(exist_ok=True)

def wait(page: Page, ms: int = 1500):
    page.wait_for_timeout(ms)

def dismiss_toasts(page: Page):
    """Cierra notificaciones/toasts que puedan tapar la UI."""
    try:
        page.locator(".o_notification_close").click(timeout=1000)
    except:
        pass

def screenshot(page: Page, name: str, wait_ms: int = 1500):
    """Toma screenshot y lo guarda."""
    wait(page, wait_ms)
    dismiss_toasts(page)
    path = SCREENSHOTS_DIR / f"{name}.png"
    page.screenshot(path=str(path), full_page=False)
    print(f"  [OK] {name}.png")
    return str(path)

def wait_for_odoo(page: Page):
    """Espera que la página de Odoo termine de cargar."""
    try:
        page.wait_for_load_state("load", timeout=30000)
    except:
        pass
    # Espera que no haya loading overlays
    try:
        page.wait_for_selector(".o_loading_indicator", state="hidden", timeout=15000)
    except:
        pass
    # Espera que el DOM principal esté listo
    try:
        page.wait_for_selector(".o_main_navbar, .o_login_form, #wrapwrap", timeout=15000)
    except:
        pass

def login(page: Page):
    """Hace login en Odoo."""
    print("Iniciando sesión...")
    page.goto(f"{URL}/odoo/login")
    wait_for_odoo(page)
    page.fill("#login", ADMIN_USER)
    page.fill("#password", ADMIN_PASS)
    page.click("button[type=submit]")
    wait_for_odoo(page)
    print("  [OK] Login exitoso")

def install_modules_via_rpc():
    """Instala los módulos necesarios vía XML-RPC."""
    print("\nInstalando módulos via XML-RPC...")
    common = xmlrpc.client.ServerProxy(f"{URL}/xmlrpc/2/common")
    uid = common.authenticate(DB, ADMIN_USER, ADMIN_PASS, {})
    models = xmlrpc.client.ServerProxy(f"{URL}/xmlrpc/2/object")

    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    # Verificar qué módulos están instalados
    installed = call("ir.module.module", "search_read",
        [[["name", "in", ["l10n_do_accounting", "l10n_do", "account"]], ["state", "=", "installed"]]],
        {"fields": ["name", "state"]})
    installed_names = {m["name"] for m in installed}
    print(f"  Módulos ya instalados: {installed_names}")

    # Instalar l10n_do_accounting si no está
    if "l10n_do_accounting" not in installed_names:
        module_ids = call("ir.module.module", "search",
            [[["name", "=", "l10n_do_accounting"]]])
        if not module_ids:
            raise Exception("Módulo l10n_do_accounting no encontrado en la instancia")
        print(f"  Instalando l10n_do_accounting (id={module_ids[0]})...")
        call("ir.module.module", "button_immediate_install", [module_ids])
        print("  [OK] Módulo instalado - esperando reinicio...")
        time.sleep(15)
    else:
        print("  [OK] l10n_do_accounting ya está instalado")

    return uid, models

def setup_company_via_rpc(uid, models):
    """Configura la compañía con RNC dominicano."""
    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    print("\nConfigurando compañía...")
    # Obtener la compañía principal
    company_ids = call("res.company", "search", [[]])
    company = call("res.company", "read", [company_ids[:1]],
        {"fields": ["name", "vat", "country_id"]})[0]
    print(f"  Compañía: {company['name']}")

    # Obtener país República Dominicana
    do_country = call("res.country", "search_read",
        [[["code", "=", "DO"]]], {"fields": ["id", "name"]})[0]

    # Actualizar si es necesario
    if not company.get("vat") or company.get("country_id", [None])[0] != do_country["id"]:
        call("res.company", "write", [company_ids[:1], {
            "vat": "101234567",
            "country_id": do_country["id"],
            "name": "Empresa Demo RD S.R.L.",
            "street": "Av. Winston Churchill #1099",
            "city": "Santo Domingo",
        }])
        print("  [OK] Compañía actualizada con RNC dominicano")
    else:
        print("  [OK] Compañía ya configurada")

    return company_ids[0]

def setup_chart_of_accounts_via_rpc(uid, models, company_id):
    """Verifica si el plan contable ya está instalado (se instala vía browser)."""
    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    print("\nVerificando plan contable...")
    accounts = call("account.account", "search_count", [[]])
    if accounts > 10:
        print(f"  [OK] Plan contable ya configurado ({accounts} cuentas)")
    else:
        print(f"  Plan contable NO instalado ({accounts} cuentas) — se instalará vía browser")

def create_contact_via_rpc(uid, models):
    """Crea contactos de prueba."""
    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    print("\nCreando contactos de prueba...")
    do_country = call("res.country", "search_read",
        [[["code", "=", "DO"]]], {"fields": ["id"]})[0]

    # Verificar si ya existen
    existing = call("res.partner", "search_count",
        [[["vat", "=", "130001901"]]])
    if existing:
        contrib_id = call("res.partner", "search",
            [[["vat", "=", "130001901"]]])[0]
        print("  [OK] Contactos ya existen")
    else:
        # Contribuyente fiscal (empresa con RNC)
        contrib_id = call("res.partner", "create", [{
            "name": "Cliente Contribuyente S.R.L.",
            "vat": "130001901",
            "country_id": do_country["id"],
            "is_company": True,
        }])
        print(f"  [OK] Contribuyente creado: id={contrib_id}")

    # Persona física (cédula)
    existing_cf = call("res.partner", "search_count",
        [[["vat", "=", "00112345678"]]])
    if not existing_cf:
        cf_id = call("res.partner", "create", [{
            "name": "Juan Pérez (Persona Física)",
            "vat": "00112345678",
            "country_id": do_country["id"],
            "is_company": False,
        }])
        print(f"  [OK] Persona física creada: id={cf_id}")

    # Proveedor extranjero
    us_country = call("res.country", "search_read",
        [[["code", "=", "US"]]], {"fields": ["id"]})[0]
    existing_ext = call("res.partner", "search_count",
        [[["name", "=", "Proveedor Extranjero LLC"]]])
    if not existing_ext:
        ext_id = call("res.partner", "create", [{
            "name": "Proveedor Extranjero LLC",
            "country_id": us_country["id"],
            "is_company": True,
        }])
        print(f"  [OK] Proveedor extranjero creado: id={ext_id}")

    return contrib_id

def get_sales_journal_id(uid, models):
    """Obtiene el ID del diario de ventas."""
    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    journals = call("account.journal", "search_read",
        [[["type", "=", "sale"]]], {"fields": ["id", "name"]})
    if journals:
        return journals[0]["id"]
    return None

def create_ncf_batch_via_rpc(uid, models, journal_id, jdt_id, doc_type_id, prefix,
                              start_num, end_num, date_from=None, date_to=None):
    """Crea un lote NCF para un tipo de documento dado."""
    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    # Verificar si ya hay un lote activo
    existing = call("l10n_do.ncf.batch", "search_count",
        [[["journal_id", "=", journal_id],
          ["l10n_latam_document_type_id", "=", doc_type_id],
          ["l10n_do_state", "=", "active"]]])
    if existing:
        print(f"  [OK] Lote {prefix} ya existe")
        return True

    # Formato de secuencia: prefijo + número con ceros
    seq_len = 10 if prefix[0] == "E" else 8
    seq_start = f"{prefix}{str(start_num).zfill(seq_len)}"
    seq_end = f"{prefix}{str(end_num).zfill(seq_len)}"

    # Crear el lote
    try:
        batch_id = call("l10n_do.ncf.batch", "create", [{
            "journal_id": journal_id,
            "journal_document_type_id": jdt_id,
            "l10n_latam_document_type_id": doc_type_id,
            "l10n_do_sequence_start": seq_start,
            "l10n_do_sequence_end": seq_end,
            "l10n_do_date_to": date_to or "2026-12-31",
            "l10n_do_date_from": date_from or "2025-01-01",
        }])
        print(f"  [OK] Lote {prefix} creado: id={batch_id} ({seq_start}-{seq_end})")
        return batch_id
    except Exception as e:
        print(f"  Error creando lote {prefix}: {e}")
        return None


def setup_ncf_batches(uid, models):
    """Configura lotes NCF para todos los diarios relevantes."""
    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    print("\nConfigurando lotes NCF...")

    sales_journal = call("account.journal", "search", [[["type", "=", "sale"]]])
    purchase_journal = call("account.journal", "search", [[["type", "=", "purchase"]]])

    if not sales_journal or not purchase_journal:
        print("  No se encontraron diarios")
        return None, None

    sj_id = sales_journal[0]
    pj_id = purchase_journal[0]

    # Lotes de ventas: (jdt_id, doc_type_id, prefix, start, end)
    # jdt_id viene de l10n_do.account.journal.document_type
    sales_batches = [
        (1, 1, "B01", 1, 500),    # Crédito Fiscal
        (2, 2, "B02", 1, 200),    # Consumo
        (3, 3, "B03", 1, 50),     # Nota de Débito
        (4, 4, "B04", 1, 100),    # Nota de Crédito
        (6, 8, "B14", 1, 100),    # Régimen Especial
    ]
    for jdt_id, dt_id, prefix, start, end in sales_batches:
        create_ncf_batch_via_rpc(uid, models, sj_id, jdt_id, dt_id, prefix, start, end)

    # Lotes de compras
    purchase_batches = [
        (16, 5, "B11", 1, 200),   # Comprobante de Compra
        (17, 7, "B13", 1, 100),   # Gasto Menor
        (18, 11, "B17", 1, 50),   # Pago al Exterior
    ]
    for jdt_id, dt_id, prefix, start, end in purchase_batches:
        create_ncf_batch_via_rpc(uid, models, pj_id, jdt_id, dt_id, prefix, start, end)

    return sj_id, pj_id


# ─── FLUJOS DE PANTALLA ───────────────────────────────────────────────────────

def install_chart_of_accounts_if_needed(uid, models):
    """Instala el plan contable dominicano via docker exec si no está instalado."""
    import subprocess

    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    accounts = call("account.account", "search_count", [[]])
    if accounts > 10:
        print(f"  [OK] Plan contable ya instalado ({accounts} cuentas)")
        return

    print("\n=== Instalando Plan Contable Dominicano vía docker exec ===")
    script = """
import sys
sys.path.insert(0, '/usr/lib/python3/dist-packages')
import odoo
from odoo.tools import config
config.parse_config([
    '--database', 'test_v19e',
    '--db_host', '172.28.57.51',
    '--db_port', '5432',
    '--db_user', 'dev',
    '--db_password', 'strong.password',
    '--addons-path', '/usr/lib/python3/dist-packages/odoo/addons,/var/lib/odoo/addons/19.0,/mnt/extra-addons,/usr/lib/python3/dist-packages/addons',
])
from odoo import api, SUPERUSER_ID
from odoo.orm.registry import Registry
registry = Registry('test_v19e')
with registry.cursor() as cr:
    env = api.Environment(cr, SUPERUSER_ID, {})
    company = env['res.company'].browse(1)
    env['account.chart.template'].try_loading('do', company, install_demo=False)
    cr.commit()
    print('OK:', env['account.account'].search_count([]), 'cuentas')
"""
    result = subprocess.run(
        ["docker", "exec", "lfernandez_localizacion_odoo", "python3", "-c", script],
        capture_output=True, text=True, timeout=120
    )
    output = result.stdout + result.stderr
    # Look for success line
    for line in output.split('\n'):
        if 'OK:' in line or 'cuentas' in line.lower():
            print(f"  {line.strip()}")

    accounts_after = call("account.account", "search_count", [[]])
    if accounts_after > 10:
        print(f"  [OK] Plan contable instalado: {accounts_after} cuentas")
    else:
        print(f"  Error: solo {accounts_after} cuentas después de la instalación")


def capture_installation(page: Page):
    """Sec 2: Instalación del módulo."""
    print("\n=== 2. Instalación del módulo ===")
    page.goto(f"{URL}/odoo/settings?modules=1")
    wait_for_odoo(page)
    wait(page, 2000)
    # Buscar el módulo
    try:
        search = page.locator(".o_searchbar_input, input[placeholder*='Search'], input[placeholder*='Buscar']").first
        search.fill("Fiscal Accounting")
        wait(page, 1500)
        screenshot(page, "02_01_module_search")

        # Si está instalado, capturar estado instalado
        installed_badge = page.locator(".o_module_install_btn:has-text('Instalado'), .o_module_install_btn:has-text('Installed')").first
        if installed_badge.count() > 0:
            screenshot(page, "02_02_module_installed")
        else:
            screenshot(page, "02_02_module_install_button")
    except Exception as e:
        print(f"  Advertencia: {e}")
        screenshot(page, "02_01_apps_screen")


def capture_company_config(page: Page):
    """Sec 3.1: Configuración de la compañía."""
    print("\n=== 3.1 Configuración de la Compañía ===")
    page.goto(f"{URL}/odoo/settings/company")
    wait_for_odoo(page)
    wait(page, 2000)
    screenshot(page, "03_01_company_form")


def capture_chart_of_accounts(page: Page):
    """Sec 3.2: Verificación del plan contable."""
    print("\n=== 3.2 Plan Contable ===")
    # Ir a Configuración de Facturación
    page.goto(f"{URL}/odoo/accounting/settings")
    wait_for_odoo(page)
    wait(page, 2000)
    screenshot(page, "03_02_accounting_settings_coa")


def capture_contact_config(page: Page, uid, models):
    """Sec 3.3: Configuración de contactos."""
    print("\n=== 3.3 Configuración de Contactos ===")

    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    # Ir a contacto contribuyente
    partner_ids = call("res.partner", "search", [[["vat", "=", "130001901"]]])
    if partner_ids:
        page.goto(f"{URL}/odoo/contacts/{partner_ids[0]}")
        wait_for_odoo(page)
        wait(page, 2000)
        screenshot(page, "03_03_contact_contribuyente")

    # Persona física
    partner_cf_ids = call("res.partner", "search", [[["vat", "=", "00112345678"]]])
    if partner_cf_ids:
        page.goto(f"{URL}/odoo/contacts/{partner_cf_ids[0]}")
        wait_for_odoo(page)
        wait(page, 2000)
        screenshot(page, "03_03_contact_persona_fisica")


def open_journal_config(page: Page, journal_name: str = "Ventas"):
    """Abre la configuración de un diario desde el Kanban."""
    page.goto(f"{URL}/odoo/accounting/journals")
    wait_for_odoo(page)
    wait(page, 2000)

    # Hover sobre la tarjeta del diario para mostrar el dropdown
    card = page.locator(".o_kanban_record").filter(has_text=journal_name).first
    card.hover()
    wait(page, 500)

    # Click en el dropdown de opciones (…)
    dropdown_btn = card.locator(".o_dropdown_kanban button").first
    dropdown_btn.click(force=True)
    wait(page, 800)

    # Click en Configuración
    config_item = page.locator(".dropdown-menu .dropdown-item").filter(has_text="Configuración").first
    config_item.click()
    wait_for_odoo(page)
    wait(page, 2000)


def capture_journal_config(page: Page, uid, models, journal_id):
    """Sec 3.4: Configuración de lotes NCF."""
    print("\n=== 3.4 Configuración de Lotes NCF ===")

    try:
        open_journal_config(page, "Ventas")
        screenshot(page, "03_04_journal_ventas")

        # Navegar a la sección de Tipos de Documento (ya está visible)
        # Hacer scroll hacia la sección COMPROBANTES FISCALES
        page.locator("text=COMPROBANTES FISCALES").scroll_into_view_if_needed()
        wait(page, 1000)
        screenshot(page, "03_04_journal_document_types_tab")

        # Click en "Lotes" de Crédito Fiscal para ver el lote
        lotes_btn = page.locator("button:has-text('Lotes'), a:has-text('Lotes')").first
        if lotes_btn.count() > 0:
            lotes_btn.click()
            wait_for_odoo(page)
            wait(page, 1500)
            screenshot(page, "03_04_ncf_batch_from_journal")

    except Exception as e:
        print(f"  Advertencia journal config: {e}")
        screenshot(page, "03_04_journal_ventas")


def capture_ncf_batches(page: Page, uid, models):
    """Sec 3.4 / 9: Vista de lotes NCF desde el diario."""
    print("\n=== Lotes NCF ===")

    try:
        open_journal_config(page, "Ventas")
        wait(page, 1500)

        # Navegar a sección de comprobantes
        try:
            page.locator("text=COMPROBANTES FISCALES").scroll_into_view_if_needed()
            wait(page, 1000)
        except:
            pass

        screenshot(page, "03_04_ncf_batch_form")

        # Click en "Lotes" de Crédito Fiscal
        lotes_btn = page.locator("button:has-text('Lotes'), a:has-text('Lotes')").first
        if lotes_btn.count() > 0:
            lotes_btn.click()
            wait_for_odoo(page)
            wait(page, 1500)
            screenshot(page, "03_04_ncf_batch_list")
        else:
            screenshot(page, "03_04_ncf_batch_list")

    except Exception as e:
        print(f"  Advertencia: {e}")
        screenshot(page, "03_04_ncf_batch_form")


def capture_invoice_creation(page: Page, uid, models):
    """Sec 4: Emisión de facturas de venta."""
    print("\n=== 4. Emisión de Facturas de Venta ===")

    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    # Buscar cliente y recursos
    partner_ids = call("res.partner", "search", [[["vat", "=", "130001901"]]])
    if not partner_ids:
        print("  No hay contribuyente")
        return None

    accounts = call("account.account", "search_read",
        [[["account_type", "=", "income"]]], {"fields": ["id"], "limit": 1})
    account_id = accounts[0]["id"] if accounts else None

    taxes = call("account.tax", "search_read",
        [[["type_tax_use", "=", "sale"], ["amount", "=", 18.0]]],
        {"fields": ["id"], "limit": 1})
    tax_ids = [[6, 0, [taxes[0]["id"]]]] if taxes else []

    # Captura 1: Mostrar el formulario de nueva factura en borrador
    # Crear un borrador via RPC
    invoice_vals = {
        "move_type": "out_invoice",
        "partner_id": partner_ids[0],
        "invoice_line_ids": [[0, 0, {
            "name": "Servicios de Consultoría",
            "quantity": 1,
            "price_unit": 5000.0,
            "account_id": account_id,
            "tax_ids": tax_ids,
        }]],
    }
    draft_id = call("account.move", "create", [invoice_vals])

    # Navegar al borrador para capturas
    page.goto(f"{URL}/odoo/accounting/customer-invoices/{draft_id}")
    wait_for_odoo(page)
    wait(page, 2500)
    screenshot(page, "04_01_invoice_new_customer")
    screenshot(page, "04_02_invoice_draft_with_line")

    # Confirmar desde la UI
    try:
        confirm_btn = page.locator("button:has-text('Confirmar'), button:has-text('Confirm')").first
        if confirm_btn.count() > 0:
            confirm_btn.click()
            wait_for_odoo(page)
            wait(page, 2500)
    except Exception as e:
        # Confirmar vía RPC como fallback
        try:
            call("account.move", "action_post", [[draft_id]])
        except:
            pass
        page.goto(f"{URL}/odoo/accounting/customer-invoices/{draft_id}")
        wait_for_odoo(page)
        wait(page, 2000)

    screenshot(page, "04_03_invoice_confirmed_with_ncf")

    # Verificar NCF asignado
    try:
        inv = call("account.move", "read", [[draft_id]],
            {"fields": ["l10n_latam_document_number", "state"]})[0]
        print(f"  NCF asignado: {inv.get('l10n_latam_document_number')} estado={inv.get('state')}")
    except:
        pass

    return draft_id


def capture_purchase_invoice(page: Page, uid, models):
    """Sec 5: Registro de facturas de compra."""
    print("\n=== 5. Facturas de Compra ===")

    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    # Escenario A: Proveedor con RNC
    page.goto(f"{URL}/odoo/accounting/vendor-bills/new")
    wait_for_odoo(page)
    wait(page, 2000)

    try:
        partner_field = page.locator("[name='partner_id'] input").first
        partner_field.fill("Cliente Contribuyente")
        wait(page, 1500)
        page.locator(".o_field_many2one_suggestion, .dropdown-item").first.click()
        wait(page, 2000)
        screenshot(page, "05_A_vendor_bill_contribuyente_ncf_field")
    except Exception as e:
        print(f"  Advertencia Escenario A: {e}")
        screenshot(page, "05_A_vendor_bill_new")

    # Escenario B: Persona física
    page.goto(f"{URL}/odoo/accounting/vendor-bills/new")
    wait_for_odoo(page)
    wait(page, 2000)

    try:
        partner_field = page.locator("[name='partner_id'] input").first
        partner_field.fill("Juan Pérez")
        wait(page, 1500)
        page.locator(".o_field_many2one_suggestion, .dropdown-item").first.click()
        wait(page, 2000)
        screenshot(page, "05_B_vendor_bill_persona_fisica")
    except Exception as e:
        print(f"  Advertencia Escenario B: {e}")
        screenshot(page, "05_B_vendor_bill_persona_fisica")


def capture_credit_note(page: Page, invoice_id, uid, models):
    """Sec 6: Nota de crédito."""
    print("\n=== 6. Notas de Crédito ===")
    if not invoice_id:
        print("  No hay factura confirmada para nota de crédito")
        return

    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    # Asegurarnos que la factura está confirmada (posted)
    inv = call("account.move", "read", [[invoice_id]], {"fields": ["state"]})[0]
    if inv["state"] != "posted":
        try:
            call("account.move", "action_post", [[invoice_id]])
        except:
            pass

    page.goto(f"{URL}/odoo/accounting/customer-invoices/{invoice_id}")
    wait_for_odoo(page)
    wait(page, 2000)
    screenshot(page, "06_01_invoice_for_credit_note")

    try:
        credit_btn = page.locator(
            "button:has-text('Nota de Crédito'), "
            "button:has-text('Credit Note'), "
            "button:has-text('Add Credit Note'), "
            "button:has-text('Revertir')"
        ).first
        if credit_btn.count() > 0:
            credit_btn.click()
            wait(page, 1500)
            screenshot(page, "06_02_credit_note_dialog")
            # Cerrar el diálogo
            try:
                page.locator("button:has-text('Cancelar'), button:has-text('Cancel'), button.btn-close").first.click()
            except:
                page.keyboard.press("Escape")
        else:
            print("  Botón de nota de crédito no encontrado")
            screenshot(page, "06_02_credit_note_dialog")
    except Exception as e:
        print(f"  Advertencia nota de crédito: {e}")


def capture_cancellation(page: Page, invoice_id):
    """Sec 8: Cancelación de facturas."""
    print("\n=== 8. Cancelación de Facturas ===")
    if not invoice_id:
        print("  No hay factura para cancelar")
        return

    page.goto(f"{URL}/odoo/accounting/customer-invoices/{invoice_id}")
    wait_for_odoo(page)
    wait(page, 2000)

    try:
        cancel_btn = page.locator("button:has-text('Cancelar'), button:has-text('Reset to Draft')").first
        if cancel_btn.count() > 0:
            cancel_btn.click()
            wait(page, 1500)
            screenshot(page, "08_01_cancel_dialog")
            # Cerrar el diálogo
            try:
                page.locator("button:has-text('Cancelar'), button:has-text('Cancel'), button.btn-close").first.click()
            except:
                page.keyboard.press("Escape")
        else:
            screenshot(page, "08_01_invoice_cancel_button")
    except Exception as e:
        print(f"  Advertencia cancelación: {e}")


def capture_ncf_batch_management(page: Page, uid, models):
    """Sec 9: Gestión de lotes NCF."""
    print("\n=== 9. Gestión de Lotes NCF ===")

    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    batches = call("l10n_do.ncf.batch", "search_read",
        [[]], {"fields": ["id", "display_name", "l10n_do_state",
                         "l10n_do_sequence_start", "l10n_do_sequence_end",
                         "l10n_do_date_to"]})

    if batches:
        # Vista del primer lote
        batch_id = batches[0]["id"]
        page.goto(f"{URL}/odoo/l10n-do-ncf-batch/{batch_id}")
        wait_for_odoo(page)
        wait(page, 2000)
        screenshot(page, "09_01_ncf_batch_detail")

        # Lista de todos los lotes
        page.goto(f"{URL}/odoo/l10n-do-ncf-batch")
        wait_for_odoo(page)
        wait(page, 2000)
        screenshot(page, "09_02_ncf_batch_list_all")

    # Vista desde el diario
    page.goto(f"{URL}/odoo/accounting/journals")
    wait_for_odoo(page)
    wait(page, 2000)
    screenshot(page, "09_03_journals_list")


def capture_invoice_print(page: Page, invoice_id):
    """Sec 11: Impresión de facturas."""
    print("\n=== 11. Impresión de Facturas ===")
    if not invoice_id:
        print("  No hay factura para imprimir")
        return

    page.goto(f"{URL}/odoo/accounting/customer-invoices/{invoice_id}")
    wait_for_odoo(page)
    wait(page, 2000)

    try:
        # Buscar botón de envío/impresión
        print_btn = page.locator("button:has-text('Enviar'), button:has-text('Send'), button:has-text('Imprimir'), button:has-text('Print')").first
        if print_btn.count() > 0:
            print_btn.click()
            wait(page, 1500)
            screenshot(page, "11_01_send_print_dialog")
            page.keyboard.press("Escape")
        else:
            screenshot(page, "11_01_invoice_for_print")
    except Exception as e:
        print(f"  Advertencia impresión: {e}")


def capture_user_permissions(page: Page):
    """Sec 12: Permisos de usuario."""
    print("\n=== 12. Permisos de Usuario ===")
    page.goto(f"{URL}/odoo/settings/users")
    wait_for_odoo(page)
    wait(page, 2000)
    screenshot(page, "12_01_users_list")

    # Abrir usuario admin para ver permisos
    try:
        admin_row = page.locator("tr.o_data_row").first
        if admin_row.count() > 0:
            admin_row.click()
            wait_for_odoo(page)
            wait(page, 2000)
            screenshot(page, "12_02_user_form_permissions")
    except Exception as e:
        print(f"  Advertencia permisos: {e}")


# ─── INSERCIÓN EN WORD ────────────────────────────────────────────────────────

def insert_screenshots_into_docx():
    """Inserta los screenshots en el documento Word."""
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from lxml import etree

    print("\n=== Insertando imágenes en el documento Word ===")

    doc_path = "doc/doc_usuario_l10n_do_accounting.docx"
    backup_path = "doc/doc_usuario_l10n_do_accounting_backup.docx"

    import shutil
    shutil.copy2(doc_path, backup_path)
    print(f"  Backup creado: {backup_path}")

    doc = Document(doc_path)

    # Mapa: texto parcial del heading/párrafo → imágenes a insertar DESPUÉS
    screenshot_map = {
        "2.2 Procedimiento de Instalación": [
            ("02_01_module_search", "Búsqueda del módulo 'Fiscal Accounting (Rep. Dominicana)' en la tienda de aplicaciones"),
            ("02_02_module_install_button", "Módulo instalado — estado confirmado en la tienda de aplicaciones"),
        ],
        "3.1 Configuración de la Compañía": [
            ("03_01_company_form", "Formulario de configuración de la compañía con RNC dominicano"),
        ],
        "3.2 Verificación del Plan Contable": [
            ("03_02_accounting_settings_coa", "Ajustes de Contabilidad mostrando el Plan Contable Dominicano activo"),
        ],
        "3.3 Configuración de Contactos": [
            ("03_03_contact_contribuyente", "Contacto tipo Contribuyente Fiscal con RNC de 9 dígitos"),
            ("03_03_contact_persona_fisica", "Contacto tipo Persona Física con cédula de 11 dígitos"),
        ],
        "3.4 Configuración de Lotes de Comprobantes Fiscales": [
            ("03_04_journal_ventas", "Diario de Ventas con la opción 'Usar documentos' activada"),
            ("03_04_journal_document_types_tab", "Sección de Comprobantes Fiscales mostrando tipos disponibles y lotes"),
            ("03_04_ncf_batch_form", "Formulario de un Lote NCF con rango de secuencias y fechas de vigencia"),
        ],
        "4. Emisión de Facturas de Venta": [
            ("04_01_invoice_new_customer", "Nueva factura con cliente seleccionado y tipo de documento asignado automáticamente"),
            ("04_02_invoice_draft_with_line", "Factura en estado borrador con línea de producto e ITBIS calculado"),
            ("04_03_invoice_confirmed_with_ncf", "Factura confirmada con el NCF asignado automáticamente"),
        ],
        "Escenario A": [
            ("05_A_vendor_bill_contribuyente_ncf_field", "Factura de proveedor con RNC — campo NCF habilitado para ingreso manual"),
        ],
        "Escenario B": [
            ("05_B_vendor_bill_persona_fisica", "Factura de proveedor informal — tipo B11 asignado automáticamente"),
        ],
        "6. Notas de Crédito Fiscales": [
            ("06_01_invoice_for_credit_note", "Factura publicada desde la cual se emitirá la nota de crédito"),
            ("06_02_credit_note_dialog", "Diálogo de creación de nota de crédito"),
        ],
        "8. Cancelación de Facturas Fiscales": [
            ("08_01_cancel_dialog", "Diálogo de cancelación con selección del Tipo de Anulación DGII"),
        ],
        "9.1 Consulta del Estado de los Lotes": [
            ("09_01_ncf_batch_detail", "Detalle de un Lote NCF con secuencias usadas, disponibles y estado"),
            ("09_02_ncf_batch_list_all", "Lista de todos los lotes NCF con su estado actual"),
        ],
        "11. Impresión de Facturas Fiscales": [
            ("11_01_send_print_dialog", "Opciones de envío e impresión de la factura fiscal"),
        ],
        "12. Administración de Permisos de Usuario": [
            ("12_01_users_list", "Lista de usuarios del sistema"),
            ("12_02_user_form_permissions", "Formulario de usuario mostrando los permisos fiscales asignables"),
        ],
    }

    def make_image_para_element(doc, img_path, caption):
        """Crea un elemento XML de párrafo con imagen, listo para insertar."""
        # Imagen
        img_para = doc.add_paragraph()
        img_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = img_para.add_run()
        run.add_picture(img_path, width=Inches(5.5))
        img_elem = img_para._element

        # Caption
        cap_para = doc.add_paragraph()
        cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_run = cap_para.add_run(caption)
        cap_run.font.size = Pt(9)
        cap_run.font.color.rgb = RGBColor(0x70, 0x70, 0x70)
        cap_run.font.italic = True
        cap_elem = cap_para._element

        # Remover de donde doc.add_paragraph() los puso (al final)
        body = doc.element.body
        body.remove(img_elem)
        body.remove(cap_elem)

        return img_elem, cap_elem

    def find_section_end_element(doc, heading_text):
        """
        Encuentra el elemento del último párrafo de una sección dado el texto del heading.
        Retorna el último elemento antes del siguiente heading.
        """
        found_heading = False
        last_elem = None

        for para in doc.paragraphs:
            if not found_heading:
                if heading_text in para.text:
                    found_heading = True
                    last_elem = para._element
            else:
                if para.style.name.startswith("Heading") and para.text.strip():
                    # Llegamos al siguiente heading, el final de la sección es el elemento anterior
                    break
                if para.text.strip() or para._element.findall(
                    './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing'
                ):
                    last_elem = para._element

        return last_elem

    # Insertar imágenes sección por sección
    inserted_sections = set()
    for target_text, images in screenshot_map.items():
        # Evitar insertar dos veces
        if target_text in inserted_sections:
            continue

        # Encontrar el último elemento de la sección
        anchor_elem = find_section_end_element(doc, target_text)

        if anchor_elem is None:
            print(f"    Sección no encontrada: {target_text[:40]}")
            continue

        inserted_sections.add(target_text)

        # Insertar las imágenes DESPUÉS del anchor_elem (en orden normal)
        current_anchor = anchor_elem
        for img_name, caption in images:
            img_path = str(SCREENSHOTS_DIR / f"{img_name}.png")
            if not Path(img_path).exists():
                print(f"    Imagen no encontrada: {img_path}")
                continue

            try:
                img_elem, cap_elem = make_image_para_element(doc, img_path, caption)

                # Insertar después del anchor actual
                current_anchor.addnext(cap_elem)
                current_anchor.addnext(img_elem)

                # El próximo anchor es el caption (para insertar la siguiente imagen después de él)
                current_anchor = cap_elem

                print(f"    Imagen insertada: {Path(img_path).name}")
            except Exception as e:
                print(f"    Error insertando {img_name}: {e}")

    doc.save(doc_path)
    print(f"\n  [OK] Documento guardado: {doc_path}")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def setup_company_vat_pre_install():
    """Configura el RNC de la compañía ANTES de instalar módulos."""
    print("\nConfigurando RNC de compañía (pre-instalación)...")
    common = xmlrpc.client.ServerProxy(f"{URL}/xmlrpc/2/common")
    uid = common.authenticate(DB, ADMIN_USER, ADMIN_PASS, {})
    models = xmlrpc.client.ServerProxy(f"{URL}/xmlrpc/2/object")

    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, ADMIN_PASS, model, method, args, kwargs or {})

    company_ids = call("res.company", "search", [[]])
    company = call("res.company", "read", [company_ids[:1]],
        {"fields": ["name", "vat", "country_id"]})[0]

    # Obtener país RD
    do_country = call("res.country", "search_read",
        [[["code", "=", "DO"]]], {"fields": ["id", "name"]})[0]

    call("res.company", "write", [company_ids[:1], {
        "vat": "101234567",
        "country_id": do_country["id"],
        "name": "Empresa Demo RD S.R.L.",
        "street": "Av. Winston Churchill #1099",
        "city": "Santo Domingo",
    }])
    print(f"  [OK] RNC '101234567' y país 'República Dominicana' configurados")


def main():
    print("=" * 60)
    print("Automatización de capturas - l10n_do_accounting")
    print("=" * 60)

    # 0. Configurar compañía ANTES de instalar (requiere RNC)
    setup_company_vat_pre_install()

    # 1. Instalar módulos vía RPC
    uid, models = install_modules_via_rpc()
    company_id = setup_company_via_rpc(uid, models)
    setup_chart_of_accounts_via_rpc(uid, models, company_id)
    # Los contactos y lotes se crean DESPUÉS de instalar el plan contable (vía browser)

    # 2. Lanzar browser y capturar pantallas
    print("\n" + "=" * 60)
    print("Iniciando capturas de pantalla...")
    print("=" * 60)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
        context = browser.new_context(
            viewport={"width": 1440, "height": 900},
            locale="es"
        )
        page = context.new_page()

        login(page)

        # Instalar plan contable si no está instalado
        install_chart_of_accounts_if_needed(uid, models)

        # Re-autenticar por si hubo reinicio
        uid2, models2 = None, None
        try:
            common2 = xmlrpc.client.ServerProxy(f"{URL}/xmlrpc/2/common")
            uid2 = common2.authenticate(DB, ADMIN_USER, ADMIN_PASS, {})
            models2 = xmlrpc.client.ServerProxy(f"{URL}/xmlrpc/2/object")
            uid, models = uid2, models2
        except:
            pass

        # Crear contactos y lotes
        create_contact_via_rpc(uid, models)
        sales_journal_id, purchase_journal_id = setup_ncf_batches(uid, models)
        journal_id = sales_journal_id

        capture_installation(page)
        capture_company_config(page)
        capture_chart_of_accounts(page)
        capture_contact_config(page, uid, models)
        if journal_id:
            capture_journal_config(page, uid, models, journal_id)
        capture_ncf_batches(page, uid, models)
        invoice_id = capture_invoice_creation(page, uid, models)
        capture_purchase_invoice(page, uid, models)
        if invoice_id:
            capture_credit_note(page, invoice_id, uid, models)
            capture_cancellation(page, invoice_id)
        capture_ncf_batch_management(page, uid, models)
        if invoice_id:
            capture_invoice_print(page, invoice_id)
        capture_user_permissions(page)

        browser.close()

    print(f"\n[OK] Screenshots guardados en: {SCREENSHOTS_DIR}")
    print(f"Total imágenes: {len(list(SCREENSHOTS_DIR.glob('*.png')))}")

    # 4. Insertar en Word
    insert_screenshots_into_docx()

    print("\n" + "=" * 60)
    print("¡Proceso completado!")
    print("=" * 60)


if __name__ == "__main__":
    main()
