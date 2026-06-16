# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError
import base64
import io
import logging
import xlsxwriter

_logger = logging.getLogger(__name__)


class AccountIvaResumenWizard(models.TransientModel):
    _name = 'account.iva.resumen.wizard'
    _description = 'Asistente de Resumen de IVA Unificado (Formulario 30)'

    company_id = fields.Many2one(
        'res.company',
        string='Compañía',
        required=True,
        default=lambda self: self.env.company
    )
    date_from = fields.Date(
        string='Fecha Desde',
        required=True,
        default=fields.Date.context_today
    )
    date_to = fields.Date(
        string='Fecha Hasta',
        required=True,
        default=fields.Date.context_today
    )

    # Campos Premium para el arrastre de saldos del mes anterior
    excedente_anterior = fields.Float(
        string='Excedente de Crédito del Mes Anterior (Casilla 38)',
        default=0.0,
        help='Monto del excedente de crédito fiscal acumulado del mes anterior.'
    )
    retenciones_acumuladas = fields.Float(
        string='Retenciones Acumuladas por Descontar (Casilla 64)',
        default=0.0,
        help='Monto de retenciones acumuladas no aplicadas del mes anterior.'
    )

    pdf_file = fields.Binary(string='Archivo PDF', readonly=True)
    pdf_filename = fields.Char(string='Nombre PDF', readonly=True)
    excel_file = fields.Binary(string='Archivo Excel', readonly=True)
    excel_filename = fields.Char(string='Nombre Excel', readonly=True)

    state = fields.Selection([
        ('draft', 'Borrador'),
        ('done', 'Generado')
    ], string='Estado', default='draft')

    # =====================================================================
    # DATA METHODS
    # =====================================================================

    def _get_resumen_data(self):
        """Obtiene y calcula todos los datos unificados para el Formulario 30 de IVA."""
        self.ensure_one()
        debitos = self._get_debitos_fiscales()
        creditos = self._get_creditos_fiscales()

        # === DÉBITOS FISCALES (Ventas) ===
        # Casilla 46 (Total Ventas y Débitos Fiscales para efectos de Determinación)
        total_base_debitos = (
            debitos['ventas_no_gravadas']
            + debitos['ventas_exportacion']
            + debitos['ventas_alicuota_general']
            + debitos['ventas_alicuota_general_adicional']
            + debitos['ventas_alicuota_reducida']
        )
        total_debitos = (
            debitos['iva_alicuota_general']
            + debitos['iva_alicuota_general_adicional']
            + debitos['iva_alicuota_reducida']
        )

        # Casilla 49 (Total Débitos Fiscales = Casilla 47 + Casilla 48)
        total_debitos_fiscales = total_debitos + debitos['ajuste_periodos_anteriores']

        # === CRÉDITOS FISCALES (Compras) ===
        # Casilla 35 (Total Compras y Créditos del Periodo)
        total_base_creditos = (
            creditos['compras_no_gravadas']
            + creditos['compras_importacion_general']
            + creditos['compras_importacion_adicional']
            + creditos['compras_importacion_reducida']
            + creditos['compras_alicuota_general']
            + creditos['compras_alicuota_general_adicional']
            + creditos['compras_alicuota_reducida']
        )
        total_creditos = (
            creditos['iva_importacion_general']
            + creditos['iva_importacion_adicional']
            + creditos['iva_importacion_reducida']
            + creditos['iva_alicuota_general']
            + creditos['iva_alicuota_general_adicional']
            + creditos['iva_alicuota_reducida']
        )

        # Casilla 71 (Total Créditos Fiscales Deducibles)
        # Por defecto es igual a los créditos del periodo
        total_creditos_deducibles = total_creditos

        # Casilla 39 (Total Créditos Fiscales = 71 + 38 + 380 - 21 - 81)
        total_creditos_fiscales = (
            total_creditos_deducibles
            + self.excedente_anterior
            + creditos['ajuste_periodos_anteriores']
            - creditos['reintegro_exportadores']
            - creditos['reintegro_exonerados']
        )

        # === AUTOLIQUIDACIÓN ===
        # Casilla 53 (Total Cuota Tributaria)
        cuota_tributaria = max(0.0, total_debitos_fiscales - total_creditos_fiscales)

        # Casilla 60 (Excedente de Crédito Fiscal para el mes siguiente)
        excedente_siguiente = max(0.0, total_creditos_fiscales - total_debitos_fiscales)

        # Casilla 78 (Sub Total Impuesto a Pagar = Cuota Tributaria - Sustituidas)
        subtotal_pagar = max(0.0, cuota_tributaria)

        # Casilla 74 (Total Retenciones = 64 + 66)
        # Casilla 66 = Retenciones del periodo (provienen de ventas/clientes)
        retenciones_periodo = debitos['iva_retenido']
        total_retenciones = self.retenciones_acumuladas + retenciones_periodo

        # Casilla 55 (Retenciones Soportadas y Descontadas en esta Declaración)
        retenciones_descontadas = min(subtotal_pagar, total_retenciones)

        # Casilla 67 (Saldo de Retenciones de IVA no Aplicado)
        saldo_retenciones_no_aplicado = total_retenciones - retenciones_descontadas

        # Casilla 56 (Sub Total Impuesto a Pagar)
        subtotal_impuesto_final = subtotal_pagar - retenciones_descontadas

        # Casilla 90 (Total a Pagar)
        total_a_pagar = subtotal_impuesto_final

        return {
            'debitos': debitos,
            'creditos': creditos,
            'total_base_debitos': total_base_debitos,
            'total_debitos': total_debitos,
            'total_debitos_fiscales': total_debitos_fiscales,
            'total_base_creditos': total_base_creditos,
            'total_creditos': total_creditos,
            'total_creditos_deducibles': total_creditos_deducibles,
            'total_creditos_fiscales': total_creditos_fiscales,
            'cuota_tributaria': cuota_tributaria,
            'excedente_siguiente': excedente_siguiente,
            'subtotal_pagar': subtotal_pagar,
            'retenciones_periodo': retenciones_periodo,
            'total_retenciones': total_retenciones,
            'retenciones_descontadas': retenciones_descontadas,
            'saldo_retenciones_no_aplicado': saldo_retenciones_no_aplicado,
            'subtotal_impuesto_final': subtotal_impuesto_final,
            'total_a_pagar': total_a_pagar,
        }

    def _get_debitos_fiscales(self):
        """Calcula débitos fiscales (ventas)."""
        moves = self.env['account.move'].search([
            ('move_type', 'in', ('out_invoice', 'out_refund')),
            ('state', '=', 'posted'),
            ('invoice_date', '>=', self.date_from),
            ('invoice_date', '<=', self.date_to),
            ('company_id', '=', self.company_id.id),
        ])

        data = {
            'ventas_no_gravadas': 0.0,
            'ventas_exportacion': 0.0,
            'ventas_alicuota_general': 0.0,
            'iva_alicuota_general': 0.0,
            'ventas_alicuota_general_adicional': 0.0,
            'iva_alicuota_general_adicional': 0.0,
            'ventas_alicuota_reducida': 0.0,
            'iva_alicuota_reducida': 0.0,
            'ajuste_periodos_anteriores': 0.0,
            'iva_retenido': 0.0,
        }

        for move in moves:
            sign = -1 if move.move_type == 'out_refund' else 1
            is_export = move.partner_id.country_id and move.partner_id.country_id.code != 'VE'

            for line in move.invoice_line_ids:
                if line.display_type != 'product':
                    continue
                base = line.price_subtotal * sign
                tax_ids = line.tax_ids

                if not tax_ids or all(t.amount == 0 for t in tax_ids):
                    if is_export:
                        data['ventas_exportacion'] += base
                    else:
                        data['ventas_no_gravadas'] += base
                    continue

                amounts = [round(t.amount, 2) for t in tax_ids]
                if 16.0 in amounts and 15.0 in amounts:
                    data['ventas_alicuota_general_adicional'] += base
                elif 16.0 in amounts:
                    data['ventas_alicuota_general'] += base
                elif 8.0 in amounts:
                    data['ventas_alicuota_reducida'] += base
                else:
                    data['ventas_no_gravadas'] += base

            for tax_line in move.line_ids.filtered(lambda l: l.display_type == 'tax'):
                tax_amount = abs(tax_line.balance) * sign
                tax_obj = tax_line.tax_line_id
                if not tax_obj:
                    continue

                # Identificar retenciones de clientes recibidas
                if any(word in tax_obj.name.lower() for word in ['retencion', 'retenido', 'wh']):
                    data['iva_retenido'] += tax_amount
                    continue

                amount = round(tax_obj.amount, 2)
                if amount == 16.0:
                    if any(round(t.amount, 2) == 15.0 for t in move.line_ids.tax_line_id):
                        data['iva_alicuota_general_adicional'] += tax_amount
                    else:
                        data['iva_alicuota_general'] += tax_amount
                elif amount == 15.0:
                    data['iva_alicuota_general_adicional'] += tax_amount
                elif amount == 8.0:
                    data['iva_alicuota_reducida'] += tax_amount

        return data

    def _get_creditos_fiscales(self):
        """Calcula créditos fiscales (compras)."""
        moves = self.env['account.move'].search([
            ('move_type', 'in', ('in_invoice', 'in_refund')),
            ('state', '=', 'posted'),
            ('invoice_date', '>=', self.date_from),
            ('invoice_date', '<=', self.date_to),
            ('company_id', '=', self.company_id.id),
        ])

        data = {
            'compras_no_gravadas': 0.0,
            'compras_importacion_general': 0.0,
            'iva_importacion_general': 0.0,
            'compras_importacion_adicional': 0.0,
            'iva_importacion_adicional': 0.0,
            'compras_importacion_reducida': 0.0,
            'iva_importacion_reducida': 0.0,
            'compras_alicuota_general': 0.0,
            'iva_alicuota_general': 0.0,
            'compras_alicuota_general_adicional': 0.0,
            'iva_alicuota_general_adicional': 0.0,
            'compras_alicuota_reducida': 0.0,
            'iva_alicuota_reducida': 0.0,
            'ajuste_periodos_anteriores': 0.0,
            'reintegro_exportadores': 0.0,
            'reintegro_exonerados': 0.0,
            'iva_retenido': 0.0,
        }

        for move in moves:
            sign = -1 if move.move_type == 'in_refund' else 1
            is_import = move.partner_id.country_id and move.partner_id.country_id.code != 'VE'

            for line in move.invoice_line_ids:
                if line.display_type != 'product':
                    continue
                base = line.price_subtotal * sign
                tax_ids = line.tax_ids

                if not tax_ids or all(t.amount == 0 for t in tax_ids):
                    data['compras_no_gravadas'] += base
                    continue

                amounts = [round(t.amount, 2) for t in tax_ids]

                if is_import:
                    if 16.0 in amounts and 15.0 in amounts:
                        data['compras_importacion_adicional'] += base
                    elif 16.0 in amounts:
                        data['compras_importacion_general'] += base
                    elif 8.0 in amounts:
                        data['compras_importacion_reducida'] += base
                    else:
                        data['compras_no_gravadas'] += base
                else:
                    if 16.0 in amounts and 15.0 in amounts:
                        data['compras_alicuota_general_adicional'] += base
                    elif 16.0 in amounts:
                        data['compras_alicuota_general'] += base
                    elif 8.0 in amounts:
                        data['compras_alicuota_reducida'] += base
                    else:
                        data['compras_no_gravadas'] += base

            for tax_line in move.line_ids.filtered(lambda l: l.display_type == 'tax'):
                tax_amount = abs(tax_line.balance) * sign
                tax_obj = tax_line.tax_line_id
                if not tax_obj:
                    continue

                if any(word in tax_obj.name.lower() for word in ['retencion', 'retenido', 'wh']):
                    data['iva_retenido'] += tax_amount
                    continue

                amt = round(tax_obj.amount, 2)
                if is_import:
                    if amt == 16.0:
                        if any(round(t.amount, 2) == 15.0 for t in move.line_ids.tax_line_id):
                            data['iva_importacion_adicional'] += tax_amount
                        else:
                            data['iva_importacion_general'] += tax_amount
                    elif amt == 15.0:
                        data['iva_importacion_adicional'] += tax_amount
                    elif amt == 8.0:
                        data['iva_importacion_reducida'] += tax_amount
                else:
                    if amt == 16.0:
                        if any(round(t.amount, 2) == 15.0 for t in move.line_ids.tax_line_id):
                            data['iva_alicuota_general_adicional'] += tax_amount
                        else:
                            data['iva_alicuota_general'] += tax_amount
                    elif amt == 15.0:
                        data['iva_alicuota_general_adicional'] += tax_amount
                    elif amt == 8.0:
                        data['iva_alicuota_reducida'] += tax_amount

        return data

    # =====================================================================
    # ACTIONS
    # =====================================================================

    def action_generate_pdf(self):
        self.ensure_one()
        data = self._get_resumen_data()
        if not data:
            raise UserError(_("No se encontró información"))

        report_action = self.env.ref('l10n_ve_simplit_fiscal.action_resumen_iva_report')
        template_id = 'l10n_ve_simplit_fiscal.resumen_iva_template'
        pdf_content, _type = report_action._render_qweb_pdf(template_id, res_ids=[self.id])

        filename = f'Resumen_IVA_{self.date_from}_{self.date_to}.pdf'
        self.write({
            'state': 'done',
            'pdf_file': base64.b64encode(pdf_content),
            'pdf_filename': filename
        })

        return {
            'type': 'ir.actions.act_window',
            'res_model': self._name,
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'new',
        }

    def action_generate_excel(self):
        self.ensure_one()
        data = self._get_resumen_data()
        if not data:
            raise UserError(_("No se encontró información"))

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        self._generate_excel_resumen(workbook, data)
        workbook.close()
        output.seek(0)

        filename = f'Resumen_IVA_{self.date_from}_{self.date_to}.xlsx'
        self.write({
            'state': 'done',
            'excel_file': base64.b64encode(output.read()),
            'excel_filename': filename
        })

        return {
            'type': 'ir.actions.act_window',
            'res_model': self._name,
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'new',
        }

    # =====================================================================
    # EXCEL GENERATION METHOD - 100% CORRESPONDING TO SENIAT FORM 30
    # =====================================================================

    def _generate_excel_resumen(self, workbook, data):
        sheet = workbook.add_worksheet('Formulario 30 - Resumen IVA')
        d = data['debitos']
        c = data['creditos']

        # Formatos
        title_fmt = workbook.add_format({
            'bold': True, 'font_size': 14, 'align': 'center',
            'font_name': 'Arial', 'bottom': 2, 'bottom_color': '#1B3A5C'
        })
        company_fmt = workbook.add_format({'bold': True, 'font_size': 11, 'font_name': 'Arial'})
        info_fmt = workbook.add_format({'font_size': 9, 'font_color': '#333333', 'font_name': 'Arial'})
        
        section_fmt = workbook.add_format({
            'bold': True, 'font_size': 10, 'bg_color': '#1B3A5C',
            'font_color': 'white', 'border': 1, 'align': 'center', 'font_name': 'Arial'
        })
        
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#E8EDF2', 'font_color': '#1B3A5C',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 9, 'font_name': 'Arial'
        })
        
        label_fmt = workbook.add_format({'border': 1, 'align': 'left', 'text_wrap': True, 'font_size': 8, 'font_name': 'Arial'})
        idx_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_size': 8, 'font_name': 'Arial'})
        box_fmt = workbook.add_format({'border': 1, 'align': 'center', 'bold': True, 'bg_color': '#F5F5F5', 'font_size': 8, 'font_name': 'Arial'})
        pct_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_size': 8, 'font_name': 'Arial'})
        num_fmt = workbook.add_format({'border': 1, 'num_format': '#,##0.00', 'font_size': 8, 'font_name': 'Arial'})
        empty_fmt = workbook.add_format({'border': 1, 'bg_color': '#FAFAFA'})

        total_label_fmt = workbook.add_format({
            'bold': True, 'border': 1, 'bg_color': '#E8EDF2', 'font_size': 8, 'font_name': 'Arial'
        })
        total_num_fmt = workbook.add_format({
            'bold': True, 'border': 1, 'num_format': '#,##0.00',
            'bg_color': '#E8EDF2', 'font_size': 8, 'font_name': 'Arial'
        })

        # Anchos de columnas
        sheet.set_column(0, 0, 4)   # Indice
        sheet.set_column(1, 1, 55)  # Descripción
        sheet.set_column(2, 2, 6)   # Casilla Base
        sheet.set_column(3, 3, 16)  # Base Imponible
        sheet.set_column(4, 4, 6)   # %
        sheet.set_column(5, 5, 6)   # Casilla Impuesto
        sheet.set_column(6, 6, 16)  # Débito/Crédito Fiscal
        sheet.set_column(7, 7, 4)   # Checker

        # Título y Encabezados de Datos
        sheet.merge_range('A1:H1', 'RESUMEN DE IVA (FORMULARIO 30)', title_fmt)
        sheet.write('B3', self.company_id.name.upper(), company_fmt)
        sheet.write('B4', f'R.I.F.: {self.company_id.vat or ""}', info_fmt)
        sheet.write('F3', 'PERIODO FISCAL:', info_fmt)
        sheet.write('F4', f'DESDE: {self.date_from.strftime("%d/%m/%Y")}', info_fmt)
        sheet.write('F5', f'HASTA: {self.date_to.strftime("%d/%m/%Y")}', info_fmt)

        row = 6

        # =====================================================================
        # DÉBITOS FISCALES
        # =====================================================================
        sheet.merge_range(row, 0, row, 7, 'DÉBITOS FISCALES', section_fmt)
        row += 1

        headers = ['#', 'Descripción', 'Cas.B', 'Base Imponible', 'Rate', 'Cas.I', 'Débito Fiscal', 'Chk']
        for col, h in enumerate(headers):
            sheet.write(row, col, h, header_fmt)
        row += 1

        deb_rows = [
            (1, 'VENTAS INTERNAS NO GRAVADAS', '40', d['ventas_no_gravadas'], '', '', 0.0, '0'),
            (2, 'VENTAS DE EXPORTACION', '41', d['ventas_exportacion'], '', '', 0.0, '0'),
            (3, 'VENTAS INTERNAS GRAVADAS POR ALICUOTA GENERAL', '42', d['ventas_alicuota_general'], '16', '43', d['iva_alicuota_general'], '7'),
            (4, 'VENTAS INTERNAS GRAVADAS POR ALICUOTA GENERAL MAS ALICUOTA ADICIONAL', '44', d['ventas_alicuota_general_adicional'], '31', '45', d['iva_alicuota_general_adicional'], '8'),
            (5, 'VENTAS INTERNAS GRAVADAS POR ALICUOTA REDUCIDA', '442', d['ventas_alicuota_reducida'], '8', '443', d['iva_alicuota_reducida'], '7'),
            (6, 'TOTAL VENTAS Y DEBITOS FISCALES PARA EFECTOS DE DETERMINACION', '46', data['total_base_debitos'], '', '47', data['total_debitos'], '3'),
            (7, 'AJUSTE A LOS DEBITOS FISCALES DE PERIODOS ANTERIORES', '', 0.0, '', '48', d['ajuste_periodos_anteriores'], '2'),
            (8, 'CERTIFICADO DE DEBITOS FISCALES EXONERADOS (RECIBIDOS DE ENTES EXONERADOS)', '', 0.0, '', '50', 0.0, '0'),
            (9, 'TOTAL DEBITOS FISCALES', '', 0.0, '', '49', data['total_debitos_fiscales'], '1'),
        ]

        for num, desc, cb, base, rate, ci, tax, chk in deb_rows:
            sheet.write(row, 0, num, idx_fmt)
            sheet.write(row, 1, desc, label_fmt)
            sheet.write(row, 2, cb, box_fmt if cb else empty_fmt)
            sheet.write(row, 3, base if cb else '', num_fmt if cb else empty_fmt)
            sheet.write(row, 4, rate, pct_fmt)
            sheet.write(row, 5, ci, box_fmt if ci else empty_fmt)
            sheet.write(row, 6, tax if ci else '', num_fmt if ci else empty_fmt)
            sheet.write(row, 7, chk, idx_fmt)
            row += 1

        row += 1

        # =====================================================================
        # CRÉDITOS FISCALES
        # =====================================================================
        sheet.merge_range(row, 0, row, 7, 'CRÉDITOS FISCALES', section_fmt)
        row += 1

        headers = ['#', 'Descripción', 'Cas.B', 'Base Imponible', 'Rate', 'Cas.I', 'Crédito Fiscal', 'Chk']
        for col, h in enumerate(headers):
            sheet.write(row, col, h, header_fmt)
        row += 1

        cred_rows = [
            (10, 'COMPRAS NO GRAVADAS Y/O SIN DERECHO A CREDITO FISCAL', '30', c['compras_no_gravadas'], '', '', 0.0, '0'),
            (11, 'IMPORTACION GRAVADA POR ALICUOTA GENERAL', '31', c['compras_importacion_general'], '16', '32', c['iva_importacion_general'], '8'),
            (12, 'IMPORTACIONES GRAVADAS POR ALICUOTA GENERAL MAS ADICIONAL', '310', c['compras_importacion_adicional'], '31', '320', c['iva_importacion_adicional'], '8'),
            (13, 'IMPORTACIONES GRAVADAS POR ALICUOTA REDUCIDA', '312', c['compras_importacion_reducida'], '8', '322', c['iva_importacion_reducida'], '7'),
            (14, 'COMPRAS INTERNAS GRAVADAS POR ALICUOTA GENERAL', '33', c['compras_alicuota_general'], '16', '34', c['iva_alicuota_general'], '6'),
            (15, 'COMPRAS INTERNAS GRAVADAS POR ALICUOTA GENERAL MAS ADICIONAL', '330', c['compras_alicuota_general_adicional'], '31', '340', c['iva_alicuota_general_adicional'], '8'),
            (16, 'COMPRAS INTERNAS GRAVADAS POR ALICUOTA REDUCIDA', '332', c['compras_alicuota_reducida'], '8', '342', c['iva_alicuota_reducida'], '7'),
            (17, 'TOTAL COMPRAS Y CREDITOS FISCALES DEL PERIODO', '35', data['total_base_creditos'], '', '36', data['total_creditos'], '4'),
            (18, 'CREDITOS FISCALES TOTALMENTE DEDUCIBLES', '', 0.0, '', '70', data['total_creditos_deducibles'], '0'),
            (19, 'CREDITOS FISCALES PRODUCTO DE LA APLICACIÓN DE LA PRORRATA', '', 0.0, '', '37', 0.0, '3'),
            (20, 'TOTAL CREDITOS FISCALES DEDUCIBLES', '', 0.0, '', '71', data['total_creditos_deducibles'], '9'),
            (21, 'EXCEDENTE DE CREDITOS FISCALES DEL MES ANTERIOR', '', 0.0, '', '38', self.excedente_anterior, '0'),
            (22, 'REINTEGRO SOLICITADO (SOLO EXPORTADORES)', '', 0.0, '', '21', c['reintegro_exportadores'], '9'),
            (23, 'REINTEGRO SOLICITADO (SOLO QUIEN ENTES EXONERADOS)', '', 0.0, '', '81', c['reintegro_exonerados'], '9'),
            (24, 'AJUSTES A LOS CREDITOS DE PERIODOS ANTERIORES', '', 0.0, '', '380', c['ajuste_periodos_anteriores'], '2'),
            (25, 'CERTIFICADO DE DEBITOS FISCALES EXONERADOS (EMITIDOS POR ENTES EXONERADOS)', '', 0.0, '', '82', 0.0, '8'),
            (26, 'TOTAL CREDITOS FISCALES', '', 0.0, '', '39', data['total_creditos_fiscales'], '1'),
        ]

        for num, desc, cb, base, rate, ci, tax, chk in cred_rows:
            sheet.write(row, 0, num, idx_fmt)
            sheet.write(row, 1, desc, label_fmt)
            sheet.write(row, 2, cb, box_fmt if cb else empty_fmt)
            sheet.write(row, 3, base if cb else '', num_fmt if cb else empty_fmt)
            sheet.write(row, 4, rate, pct_fmt)
            sheet.write(row, 5, ci, box_fmt if ci else empty_fmt)
            sheet.write(row, 6, tax if ci else '', num_fmt if ci else empty_fmt)
            sheet.write(row, 7, chk, idx_fmt)
            row += 1

        row += 1

        # =====================================================================
        # AUTOLIQUIDACIÓN
        # =====================================================================
        sheet.merge_range(row, 0, row, 7, 'AUTOLIQUIDACIÓN', section_fmt)
        row += 1

        headers = ['#', 'Descripción', '', '', '', 'Casilla', 'Monto', 'Chk']
        for col, h in enumerate(headers):
            sheet.write(row, col, h, header_fmt)
        row += 1

        auto_rows = [
            (27, 'TOTAL CUOTA TRIBUTARIA', '53', data['cuota_tributaria'], '7'),
            (28, 'EXCEDENTE DE CREDITO FISCAL PARA EL MES SIGUIENTE', '60', data['excedente_siguiente'], '0'),
            (29, 'IMPUESTO PAGADO EN DECLARACION SUSTITUIDA', '22', 0.0, '8'),
            (30, 'RETENCIONES DESCONTADAS EN DECLARACION SUSTITUIDA', '61', 0.0, '8'),
            (31, 'PERCEPCIONES DESCONTADAS EN DECLARACION SUSTITUIDA', '24', 0.0, '6'),
            (32, 'SUB TOTAL IMPUESTO A PAGAR', '78', data['subtotal_pagar'], '2'),
            (33, 'RETENCIONES ACUMULADAS POR DESCONTAR', '64', self.retenciones_acumuladas, '6'),
            (34, 'RETENCIONES DEL PERIODO', '66', data['retenciones_periodo'], '4'),
            (35, 'CREDITOS ADQUIRIDOS POR CESION DE RETENCIONES', '72', 0.0, '8'),
            (36, 'RECUPERACION DE RETENCIONES SOLICITADO', '73', 0.0, '7'),
            (37, 'TOTAL RETENCIONES', '74', data['total_retenciones'], '6'),
            (38, 'RETENCIONES SOPORTADAS Y DESCONTADAS EN ESTA DECLARACION', '55', data['retenciones_descontadas'], '5'),
            (39, 'SALDO DE RETENCIONES DE IVA NO APLICADO', '67', data['saldo_retenciones_no_aplicado'], '3'),
            (40, 'SUB - TOTAL IMPUESTO A PAGAR', '56', data['subtotal_impuesto_final'], '4'),
            (41, 'PERCEPCIONES ACUMULADAS POR DESCONTAR', '67', 0.0, '3'),
            (42, 'PERCEPCIONES DEL PERIODO', '68', 0.0, '2'),
            (43, 'CREDITOS ADQUIRIDOS POR CESION DE PERCEPCIONES', '75', 0.0, '5'),
            (44, 'RECUPERACION DE PERCEPCIONES SOLICITADO', '76', 0.0, '4'),
            (45, 'TOTAL PERCEPCIONES', '77', 0.0, '3'),
            (46, 'PERCEPCIONES EN ADUANAS DESCONTADAS EN ESTA DECLARACION', '58', 0.0, '2'),
            (47, 'SALDO DE PERCEPCIONES EN ADUANAS NO APLICADO', '-', 0.0, '1'),
            (48, 'TOTAL A PAGAR', '90', data['total_a_pagar'], '0'),
        ]

        for num, desc, cas, val, chk in auto_rows:
            sheet.write(row, 0, num, idx_fmt)
            sheet.write(row, 1, desc, label_fmt)
            sheet.write(row, 2, '', empty_fmt)
            sheet.write(row, 3, '', empty_fmt)
            sheet.write(row, 4, '', empty_fmt)
            sheet.write(row, 5, cas, box_fmt)
            sheet.write(row, 6, val, total_num_fmt if num in [27, 28, 32, 37, 38, 40, 48] else num_fmt)
            sheet.write(row, 7, chk, idx_fmt)
            row += 1
