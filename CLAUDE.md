# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Dominican Republic (DO) localization modules for Odoo 18.0. Implements fiscal compliance with DGII (Dirección General de Impuestos Internos) regulations, including NCF/e-CF fiscal number management, tax reporting, and RNC/NCF validation against DGII services.

## Architecture

This is a **multi-module Odoo addon repository**. Each subdirectory is a separate Odoo module managed as a **git submodule** (see `.gitmodules`). The entire repo is mounted as the Odoo addons path (`/mnt/extra-addons`) inside Docker.

### Modules and Dependencies

- **l10n_do_accounting** — Core fiscal accounting module. Manages NCF/e-CF sequences, fiscal document types, income/expense types, cancellation workflows, and invoice reports. Depends on `l10n_latam_invoice_document` and `l10n_do` (Odoo core).
- **dgii_reports** — DGII tax declaration reports (606, 607, 608, 609). Depends on `l10n_do_accounting`.
- **dgii_rnc_validation** — Validates partner RNC (tax ID) numbers against DGII. Depends on `base`, `contacts`.
- **ncf_dgii_validation** — Validates NCF/e-NCF fiscal numbers against DGII. Depends on `base`, `account`.

Dependency chain: `dgii_reports` → `l10n_do_accounting` → `l10n_latam_invoice_document` + `l10n_do`

### Key Domain Concepts

- **NCF** (Número de Comprobante Fiscal): Physical fiscal invoice numbers (e.g., `B0100000001`)
- **e-CF** (Comprobante Fiscal Electrónico): Electronic fiscal numbers (e.g., `E310000000001`)
- **DGII**: Dominican Republic tax authority
- **RNC**: Tax identification number for companies
- Fiscal document types determine NCF prefix and validation rules (fiscal, consumer, informal, special, governmental, export, etc.)

## Development Environment

### Running Odoo

```bash
docker compose up        # Start Odoo on port defined in .env (default 8092)
docker compose down      # Stop
```

Odoo runs via a custom `odoo_deb:18.0` Docker image. Config is in `conf/odoo.conf`. Environment variables are in `.env`.

### Running Tests

```bash
# Run tests for a specific module inside the container
docker compose exec odoo odoo --test-enable --stop-after-init -d <database> -i <module_name>

# Example: test l10n_do_accounting
docker compose exec odoo odoo --test-enable --stop-after-init -d <database> -i l10n_do_accounting
```

Tests use `AccountTestInvoicingCommon` as base class. Test common setup is in `l10n_do_accounting/tests/common.py` with `L10nDOTestsCommon` providing DO-specific fixtures (partners, journals, document types).

### Installing/Updating Modules

```bash
docker compose exec odoo odoo -d <database> -i <module_name>     # Install
docker compose exec odoo odoo -d <database> -u <module_name>     # Update
```

## Code Conventions

- All modules follow standard Odoo 18.0 module structure: `__manifest__.py`, `__init__.py`, `models/`, `views/`, `wizard/`, `security/`, `data/`, `tests/`
- Models extend Odoo core via `_inherit` (e.g., `account.move`, `res.partner`, `account.journal`)
- Dominican-specific fields use the `l10n_do_` prefix
- Fiscal sequence logic is in `l10n_do_accounting/models/account_move.py` — the `_set_next_sequence` / `_get_last_sequence` override is the most complex part, using `is_l10n_do_seq` context flag to switch between standard Odoo sequencing and DO fiscal sequencing
- Python dependency: `pycountry` (see `requirements.txt`)
