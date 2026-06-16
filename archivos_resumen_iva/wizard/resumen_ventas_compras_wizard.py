# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError
import base64
import logging
import io
try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None

_logger = logging.getLogger(__name__)


class ResumenVentasComprasWizard(models.TransientModel):
    _name = 'resumen.ventas.compras.wizard'
    _description = 'Resumen Fiscal: Libro de Ventas y Compras'

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

    pdf_file = fields.Binary(string='Archivo PDF', readonly=True)
    pdf_filename = fields.Char(string='Nombre PDF', readonly=True)

    xls_file = fields.Binary(string='Archivo Excel', readonly=True)
    xls_filename = fields.Char(string='Nombre Excel', readonly=True)

    state = fields.Selection([
        ('draft', 'Borrador'),
        ('done', 'Generado'),
    ], string='Estado', default='draft')

    # ---------------------------------------------------------------
    # Ventas – datos detallados (una línea por factura)
    # ---------------------------------------------------------------
    def _get_sales_data(self):
        """Retorna lista de diccionarios con el detalle del Libro de Ventas."""
        Move = self.env['account.move']
        domain = [
            ('move_type', 'in', ('out_invoice', 'out_refund')),
            ('state', '=', 'posted'),
            ('invoice_date', '>=', self.date_from),
            ('invoice_date', '<=', self.date_to),
            ('company_id', '=', self.company_id.id),
        ]
        moves = Move.search(domain, order='invoice_date asc, name asc')

        lines = []
        count = 1
        for move in moves:
            sign = -1 if move.move_type == 'out_refund' else 1
            is_contribuyente = bool(move.partner_id.vat and len(move.partner_id.vat) > 5)

            base_contrib = base_no_contrib = 0.0
            iva_contrib = iva_no_contrib = 0.0
            exempt = 0.0

            for line in move.invoice_line_ids:
                iva_taxes = line.tax_ids.filtered(lambda t: t.amount > 0)
                price_subtotal = getattr(
                    line, 'l10n_ve_ta_multicurrency_price_subtotal', line.price_subtotal
                )
                if iva_taxes:
                    if is_contribuyente:
                        base_contrib += price_subtotal * sign
                    else:
                        base_no_contrib += price_subtotal * sign
                else:
                    exempt += price_subtotal * sign

            factor = move._get_ves_factor() if hasattr(move, '_get_ves_factor') else 1.0
            for tax_line in move.line_ids.filtered(lambda l: l.display_type == 'tax'):
                if tax_line.tax_line_id and tax_line.tax_line_id.amount > 0:
                    iva = (abs(tax_line.balance) * factor) * sign
                    if is_contribuyente:
                        iva_contrib += iva
                    else:
                        iva_no_contrib += iva

            wh_iva = self.env['account.wh.iva'].search(
                [('move_id', '=', move.id), ('state', '!=', 'cancel')], limit=1
            )
            iva_retenido = (
                getattr(wh_iva, 'l10n_ve_ta_multicurrency_amount_total_ret',
                        wh_iva.amount_total_ret) if wh_iva else 0.0
            ) * sign

            lines.append({
                'ope': count,
                'date': move.invoice_date,
                'rif': move.partner_id.vat or 'N/A',
                'name': move.partner_id.name or '',
                'tipo_doc': '01-FAC' if move.move_type == 'out_invoice' else '03-NC',
                'serial': move.l10n_ve_fiscal_printer_serial or '',
                'doc_num': move.l10n_ve_fiscal_invoice_number or move.name,
                'tipo_tran': '01-REG' if move.move_type == 'out_invoice' else '03-ANUL',
                'fac_afectada': (
                    move.reversed_entry_id.l10n_ve_fiscal_invoice_number
                    or move.reversed_entry_id.name
                ) if move.reversed_entry_id else '',
                'total_c_iva': getattr(
                    move, 'l10n_ve_ta_multicurrency_total_amount', move.amount_total
                ) * sign,
                'exempt': exempt,
                'base_contrib': base_contrib,
                'iva_contrib': iva_contrib,
                'base_no_contrib': base_no_contrib,
                'iva_no_contrib': iva_no_contrib,
                'iva_retenido': iva_retenido,
                'nro_ret': wh_iva.name if wh_iva else '',
                'fecha_ret': wh_iva.date if wh_iva else '',
            })
            count += 1

        return lines

    # ---------------------------------------------------------------
    # Compras – datos detallados (una línea por factura)
    # ---------------------------------------------------------------
    def _get_purchase_data(self):
        """Retorna lista de diccionarios con el detalle del Libro de Compras."""
        Move = self.env['account.move']
        domain = [
            ('move_type', 'in', ('in_invoice', 'in_refund')),
            ('state', '=', 'posted'),
            ('invoice_date', '>=', self.date_from),
            ('invoice_date', '<=', self.date_to),
            ('company_id', '=', self.company_id.id),
        ]
        moves = Move.search(domain, order='invoice_date asc, name asc')

        lines = []
        count = 1
        for move in moves:
            doc_type = 'FAC'
            if move.move_type == 'in_refund':
                doc_type = 'NC'
            elif getattr(move, 'debit_origin_id', False):
                doc_type = 'ND'

            tran_type = '01-REG'
            sign = -1 if move.move_type == 'in_refund' else 1

            wh_iva = self.env['account.wh.iva'].search(
                [('move_id', '=', move.id), ('state', '!=', 'cancel')], limit=1
            )

            mc_total = getattr(move, 'l10n_ve_ta_multicurrency_total_amount', False)
            if mc_total is not False:
                total_with_iva = move.l10n_ve_ta_multicurrency_total_amount * sign
                taxable_base = move.l10n_ve_ta_multicurrency_taxable_amount * sign
                exempt_amount = move.l10n_ve_ta_multicurrency_exempt_amount * sign
                iva_amount = move.l10n_ve_ta_multicurrency_tax_amount * sign
                wh_amount = (
                    wh_iva.l10n_ve_ta_multicurrency_amount_total_ret if wh_iva else 0.0
                ) * sign
            else:
                total_with_iva = (
                    wh_iva.amount_total_invoice if wh_iva else move.amount_total
                ) * sign
                iva_amount = 0.0
                for line in move.line_ids.filtered(lambda l: l.tax_line_id):
                    if line.tax_line_id.amount > 0:
                        iva_amount += abs(line.balance)
                iva_amount *= sign
                if wh_iva:
                    taxable_base = wh_iva.amount_taxable_base * sign
                    exempt_amount = wh_iva.amount_exempt * sign
                else:
                    taxable_base = move.amount_untaxed * sign
                    exempt_amount = 0.0
                wh_amount = (wh_iva.amount_total_ret if wh_iva else 0.0) * sign

            lines.append({
                'ope': count,
                'date': move.invoice_date,
                'rif': move.partner_id.vat or '',
                'partner_name': move.partner_id.name or '',
                'doc_number': move.l10n_ve_supplier_invoice_number or move.name,
                'doc_type': doc_type,
                'control_number': move.l10n_ve_control_number or '',
                'tran_type': tran_type,
                'affected_doc': (
                    move.reversed_entry_id.l10n_ve_supplier_invoice_number
                    or move.reversed_entry_id.name
                ) if move.reversed_entry_id else '',
                'total_with_iva': total_with_iva,
                'exempt_amount': exempt_amount,
                'taxable_base': taxable_base,
                'iva_rate': 16,
                'iva_amount': iva_amount,
                'wh_amount': wh_amount,
                'wh_rate': wh_iva.retention_percentage if wh_iva else 0.0,
                'wh_number': wh_iva.name if wh_iva else '',
                'wh_date': wh_iva.date if wh_iva else False,
            })
            count += 1

        return lines

    # ---------------------------------------------------------------
    # Acción: generar PDF combinado
    # ---------------------------------------------------------------
    def action_generate_pdf(self):
        self.ensure_one()
        if self.date_from > self.date_to:
            raise UserError(_("La fecha inicial no puede ser mayor a la fecha final."))

        report_action = self.env.ref(
            'l10n_ve_simplit_fiscal.action_report_resumen_ventas_compras'
        )
        pdf_content, _rtype = report_action._render_qweb_pdf(
            'l10n_ve_simplit_fiscal.report_resumen_ventas_compras_template',
            res_ids=[self.id],
        )

        filename = f'Resumen_Fiscal_{self.date_from}_{self.date_to}.pdf'
        self.write({
            'state': 'done',
            'pdf_file': base64.b64encode(pdf_content),
            'pdf_filename': filename,
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
        if not xlsxwriter:
            raise UserError(_("La librería 'xlsxwriter' no está instalada. Por favor, instálela para usar esta función."))

        if self.date_from > self.date_to:
            raise UserError(_("La fecha inicial no puede ser mayor a la fecha final."))

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        
        # Formatos
        title_format = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 12,
        })
        header_format = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 9,
            'border': 1
        })
        cell_format = workbook.add_format({'font_size': 8, 'border': 1, 'valign': 'vcenter'})
        cell_center = workbook.add_format({'font_size': 8, 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        num_format = workbook.add_format({'font_size': 8, 'border': 1, 'num_format': '#,##0.00', 'align': 'center', 'valign': 'vcenter'})
        bold_num_format = workbook.add_format({'bold': True, 'font_size': 8, 'border': 1, 'num_format': '#,##0.00', 'align': 'center', 'valign': 'vcenter'})
        bold_header = workbook.add_format({'bold': True, 'font_size': 8, 'border': 1, 'align': 'left', 'valign': 'vcenter'})

        # Hoja de Resumen
        sheet = workbook.add_worksheet("Resumen de IVA")
        sheet.set_column('A:A', 5)
        sheet.set_column('B:B', 60)
        sheet.set_column('C:C', 6)
        sheet.set_column('D:D', 20)
        sheet.set_column('E:E', 6)
        sheet.set_column('F:F', 6)
        sheet.set_column('G:G', 20)
        sheet.set_column('H:H', 6)

        # Encabezado
        sheet.merge_range('A1:H1', "RESUMEN DE IVA", title_format)
        
        company_name = self.company_id.name or ''
        rif = self.company_id.vat or ''
        sheet.write('A3', f"{company_name}", workbook.add_format({'bold': True, 'font_size': 10}))
        sheet.write('A4', f"{rif}", workbook.add_format({'bold': True, 'font_size': 10}))
        
        sheet.write('A5', "PERIODO FISCAL:", workbook.add_format({'bold': True, 'font_size': 10}))
        sheet.write('C5', self.date_from.strftime('%d/%m/%Y') if self.date_from else '', workbook.add_format({'bold': True, 'font_size': 10, 'align': 'center'}))
        sheet.write('C6', self.date_to.strftime('%d/%m/%Y') if self.date_to else '', workbook.add_format({'bold': True, 'font_size': 10, 'align': 'center'}))

        sales_data = self._get_sales_data()
        purchase_data = self._get_purchase_data()

        # Acumuladores
        t_ventas_base_c = sum(l.get('base_contrib', 0.0) for l in sales_data)
        t_ventas_iva_c = sum(l.get('iva_contrib', 0.0) for l in sales_data)
        t_ventas_base_nc = sum(l.get('base_no_contrib', 0.0) for l in sales_data)
        t_ventas_iva_nc = sum(l.get('iva_no_contrib', 0.0) for l in sales_data)
        t_ventas_exentas = sum(l.get('exempt', 0.0) for l in sales_data)
        t_ventas_export = sum(l.get('export_base', 0.0) for l in sales_data) if any('export_base' in l for l in sales_data) else 0.0
        
        t_compras_base = sum(l.get('taxable_base', 0.0) for l in purchase_data)
        t_compras_iva = sum(l.get('iva_amount', 0.0) for l in purchase_data)
        t_compras_exentas = sum(l.get('exempt_amount', 0.0) for l in purchase_data)
        t_retenciones_iva = sum(l.get('wh_amount', 0.0) for l in purchase_data)

        total_ventas_gravadas_base = t_ventas_base_c + t_ventas_base_nc
        total_ventas_gravadas_iva = t_ventas_iva_c + t_ventas_iva_nc
        
        total_compras_gravadas_base = t_compras_base
        total_compras_gravadas_iva = t_compras_iva
        
        impuesto_a_pagar = max(total_ventas_gravadas_iva - total_compras_gravadas_iva, 0.0)

        # Estructura de la tabla (Nro, Descripcion, Cas.Base, Val.Base, Dig.Base, Cas.Imp, Val.Imp, Dig.Imp)
        data_rows = [
            # DÉBITOS FISCALES
            ({'merge': 'A{r}:B{r}', 'val': 'DEBITOS FISCALES', 'fmt': header_format}, {'col': 2, 'merge': 'C{r}:E{r}', 'val': 'Base Imponible', 'fmt': header_format}, {'col': 5, 'merge': 'F{r}:H{r}', 'val': 'Debito Fiscal', 'fmt': header_format}),
            (1, 'VENTAS INTERNAS NO GRAVADAS', 40, t_ventas_exentas, 0, '', '', ''),
            (2, 'VENTAS DE EXPORTACION', 41, t_ventas_export, 9, '', '', ''),
            (3, 'VENTAS INTERNAS GRAVADAS POR ALICUOTA GENERAL', 42, total_ventas_gravadas_base, 8, 43, total_ventas_gravadas_iva, 7),
            (4, 'VENTAS INTERNAS GRAVADAS POR ALICUOTA GENERAL MAS ALICUOTA ADICIONAL', 442, 0.0, 8, 452, 0.0, 8),
            (5, 'VENTAS INTERNAS GRAVADAS POR ALICUOTA REDUCIDA', 443, 0.0, 7, 453, 0.0, 7),
            (6, 'TOTAL VENTAS Y DEBITOS FISCALES PARA EFECTOS DE DETERMINACION', 46, t_ventas_exentas + t_ventas_export + total_ventas_gravadas_base, 4, 47, total_ventas_gravadas_iva, 3),
            (7, 'AJUSTE A LOS DEBITOS FISCALES DE PERIODOS ANTERIORES', '', '', '', 48, 0.0, 2),
            (8, 'CERTIFICADO DE DEBITOS FISCALES EXONERADOS (RECIBIDOS DE ENTES EXONERADOS)', '', '', '', 80, 0.0, 0),
            (9, 'TOTAL DEBITOS FISCALES .... Realice la operación (47 + - 48)', '', '', '', 49, total_ventas_gravadas_iva, 1),
            
            # CRÉDITOS FISCALES
            ({'merge': 'A{r}:B{r}', 'val': 'CREDITOS FISCALES', 'fmt': header_format}, {'col': 2, 'merge': 'C{r}:E{r}', 'val': 'Base Imponible', 'fmt': header_format}, {'col': 5, 'merge': 'F{r}:H{r}', 'val': 'Credito Fiscal', 'fmt': header_format}),
            (10, 'COMPRAS NO GRAVADAS Y/O SIN DERECHO A CREDITO FISCAL', 30, t_compras_exentas, 0, '', '', ''),
            (11, 'IMPORTACION GRAVADA POR ALICUOTA GENERAL', 31, 0.0, 9, 32, 0.0, 8),
            (12, 'IMPORTACIONES GRAVADAS POR ALICUOTA GENERAL MAS ADICIONAL', 312, 0.0, 8, 322, 0.0, 8),
            (13, 'IMPORTACIONES GRAVADAS POR ALICUOTA REDUCIDA', 313, 0.0, 7, 323, 0.0, 7),
            (14, 'COMPRAS INTERNAS GRAVADAS POR ALICUOTA GENERAL', 33, total_compras_gravadas_base, 7, 34, total_compras_gravadas_iva, 6),
            (15, 'COMPRAS INTERNAS GRAVADAS POR ALICUOTA GENERAL MAS ADICIONAL', 332, 0.0, 8, 342, 0.0, 8),
            (16, 'COMPRAS INTERNAS GRAVADAS POR ALICUOTA REDUCIDA', 333, 0.0, 7, 343, 0.0, 7),
            (17, 'TOTAL COMPRAS Y CREDITOS FISCALES DEL PERIODO', 35, t_compras_exentas + total_compras_gravadas_base, 5, 36, total_compras_gravadas_iva, 4),
            (18, 'CREDITOS FISCALES TOTALMENTE DEDUCIBLES', '', '', '', 70, '', 0),
            (19, 'CREDITOS FISCALES PRODUCTO DE LA APLICACIÓN DE LA PRORRATA', '', '', '', 37, '', 3),
            (20, 'TOTAL CREDITOS FISCALES DEDUCIBLES', '', '', '', 71, '-', 9),
            (21, 'EXCEDENTE DE CREDITOS FISCALES DEL MES ANTERIOR', '', '', '', 20, 0.0, 0),
            (22, 'REINTEGRO SOLICITADO (SOLO EXPORTADORES)', '', '', '', 21, '', 9),
            (23, 'REINTEGRO SOLICITADO (SOLO QUIEN SUMINISTRE BIENES A ENTES EXONERADOS)', '', '', '', 81, '', 9),
            (24, 'AJUSTES A LOS CREDITOS DE PERIODOS ANTERIORES', '', '', '', 38, '', 2),
            (25, 'CERTIFICADO DE DEBITOS FISCALES EXONERADOS (EMITIDOS POR ENTES EXONERADOS)', '', '', '', 82, '', 8),
            (26, 'TOTAL CREDITOS FISCALES ... Realice la operación', '', '', '', 39, total_compras_gravadas_iva, 1),

            # AUTOLIQUIDACIÓN
            ({'merge': 'A{r}:H{r}', 'val': 'AUTOLIQUIDACION', 'fmt': header_format},),
            (27, 'TOTAL CUOTA TRIBUTARIA', '', '', '', 53, impuesto_a_pagar, 7),
            (28, 'EXCEDENTE DE CREDITO FISCAL PARA EL MES SIGUIENTE', '', '', '', 60, max(total_compras_gravadas_iva - total_ventas_gravadas_iva, 0.0), 0),
            (29, 'IMPUESTO PAGADO EN DECLARACION SUSTITUIDA', 22, 0.0, 8, '', '', ''),
            (30, 'RETENCIONES DESCONTADAS EN DECLARACION SUSTITUIDA', 51, 0.0, 9, '', '', ''),
            (31, 'PERCEPCIONES DESCONTADAS EN DECLARACION SUSTITUIDA', 24, 0.0, 6, '', '', ''),
            (32, 'SUB TOTAL IMPUESTO A PAGAR', '', '', '', 78, impuesto_a_pagar, 2),
            (33, 'RETENCIONES ACUMULADAS POR DESCONTAR', 54, 0.0, 6, '', '', ''),
            (34, 'RETENCIONES DEL PERIODO', 66, t_retenciones_iva, 4, '', '', ''),
            (35, 'CREDITOS ADQUIRIDOS POR CESION DE RETENCIONES', 72, '-', 8, '', '', ''),
            (36, 'RECUPERACION DE RETENCIONES SOLICITADO', 73, 0.0, 7, '', '', ''),
            (37, 'TOTAL RETENCIONES', 74, t_retenciones_iva, 6, '', '', ''),
            (38, 'RETENCIONES SOPORTADAS Y DESCONTADAS EN ESTA DECLARACION', '', '', '', 55, 0.0, 5),
            (39, 'SALDO DE RETENCIONES DE IVA NO APLICADO', 67, t_retenciones_iva, 3, '', '', ''),
            (40, 'SUB - TOTAL IMPUESTO A PAGAR', '', '', '', 56, impuesto_a_pagar, 4),
            (41, 'PERCEPCIONES ACUMULADAS EN IMPORTACIONES POR DESCONTAR', 57, '', 3, '', '', ''),
            (42, 'PERCEPCIONES DEL PERIODO', 68, '', 2, '', '', ''),
            (43, 'CREDITOS ADQUIRIDOS POR CESION DE PERCEPCIONES', 75, '', 5, '', '', ''),
            (44, 'RECUPERACION DE PERCEPCIONES SOLICITADO', 76, '', 4, '', '', ''),
            (45, 'TOTAL PERCEPCIONES', 77, '-', 3, '', '', ''),
            (0, 'PERCEPCIONES EN ADUANAS DESCONTADAS EN ESTA DECLARACION', '', '', '', 58, '-', 2),
            (47, 'SALDO DE PERCEPCIONES EN ADUANAS NO APLICADO', '', '-', 1, '', '', ''),
            (48, 'TOTAL A PAGAR ( 61 + 62 + 65 )', '', '', '', 90, impuesto_a_pagar, 0),
        ]

        row_idx = 7
        for r_data in data_rows:
            if isinstance(r_data[0], dict):
                # Header row with merged cells
                for merge_info in r_data:
                    m_range = merge_info['merge'].format(r=row_idx+1)
                    sheet.merge_range(m_range, merge_info['val'], merge_info['fmt'])
            else:
                # Normal data row
                sheet.write(row_idx, 0, r_data[0], cell_center if str(r_data[0]) else cell_format)
                sheet.write(row_idx, 1, r_data[1], cell_format)
                sheet.write(row_idx, 2, r_data[2], cell_center)
                if isinstance(r_data[3], float):
                    sheet.write(row_idx, 3, r_data[3], bold_num_format if r_data[0] in (6, 9, 17, 26) else num_format)
                else:
                    sheet.write(row_idx, 3, r_data[3], cell_center)
                sheet.write(row_idx, 4, r_data[4], cell_center)
                
                sheet.write(row_idx, 5, r_data[5], cell_center)
                if isinstance(r_data[6], float):
                    sheet.write(row_idx, 6, r_data[6], bold_num_format if r_data[0] in (6, 9, 17, 26, 32, 40, 48) else num_format)
                else:
                    sheet.write(row_idx, 6, r_data[6], cell_center)
                sheet.write(row_idx, 7, r_data[7], cell_center)
                
            row_idx += 1

        workbook.close()
        output.seek(0)
        
        filename = f'Resumen_Fiscal_{self.date_from}_{self.date_to}.xlsx'
        self.write({
            'state': 'done',
            'xls_file': base64.b64encode(output.read()),
            'xls_filename': filename,
        })

        return {
            'type': 'ir.actions.act_window',
            'res_model': self._name,
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'new',
        }
