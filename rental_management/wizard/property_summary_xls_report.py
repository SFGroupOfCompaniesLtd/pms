# -*- coding: utf-8 -*-
# Copyright 2020-Today TechKhedut.
# Part of TechKhedut. See LICENSE file for full copyright and licensing details.
from odoo import fields, models
import xlwt
import base64
from io import BytesIO


class PropertySummaryXlsReport(models.TransientModel):
    _name = 'property.summary.report.wizard'
    _description = 'Create Property Report'
    _rec_name = 'type'

    type = fields.Selection(
        [('tenancy', 'Rent'), ('sold', 'Property Sold')], string="Report For")
    start_date = fields.Date(string="Start Date" )
    end_date = fields.Date(string="End Date")
    property_ids = fields.Many2many('property.project')



    def action_property_summary_xls_report(self):
        workbook = xlwt.Workbook(encoding='utf-8')
        sheet1 = workbook.add_sheet(
            'Summary', cell_overwrite_ok=True)
        domain = [("start_date", ">=", self.start_date),
                  ("start_date", "<=", self.end_date)]
        record = self.env["property.project"].search([('id', 'in', self.property_ids.ids)])
        self.action_create_rent_contract_summary_report(
            sheet=sheet1, record=record, workbook=workbook)

        stream = BytesIO()
        workbook.save(stream)
        out = base64.encodebytes(stream.getvalue())

        attachment = self.env['ir.attachment'].sudo()
        filename = 'Rent Details' + ".xls"
        attachment_id = attachment.create(
            {'name': filename,
             'type': 'binary',
             'public': False,
             'datas': out})
        if attachment_id:
            report = {
                'type': 'ir.actions.act_url',
                'url': '/web/content/%s?download=true' % (attachment_id.id),
                'target': 'self',
            }
            return report

    def get_rent_stage(self, status):
        name = ""
        if status == "running_contract":
            name = "Running"
        elif status == "cancel_contract":
            name = "Cancel"
        elif status == "close_contract":
            name = "Close"
        elif status == "expire_contract":
            name = "Expire"
        else:
            name = "Draft"
        return name

    def get_property_type(self, type):
        name = ""
        if type == "residential":
            name = "Residential"
        elif type == "industrial":
            name = "Industrial"
        elif type == "commercial":
            name = "Commercial"
        elif type == "land":
            name = "Land"
        return name

    def get_measure_unit(self, measure_unit):
        unit = ""
        if measure_unit == "sq_ft":
            unit = "ft²"
        elif measure_unit == "sq_m":
            unit = "m²"
        elif measure_unit == "sq_yd":
            unit = "yd²"
        elif measure_unit == "cu_ft":
            unit = "ft³"
        else:
            unit = "m³"
        return unit

    def get_status(self, stage):
        name = ""
        if stage == "booked":
            name = "Booked"
        elif stage == "refund":
            name = "Refund"
        elif stage == "sold":
            name = "Sold"
        elif stage == "cancel":
            name = "Cancel"
        elif stage == "locked":
            name = "Locked"
        return name

    def get_payment_term(self, term):
        name = ""
        if term == "monthly":
            name = "Monthly"
        elif term == "full_payment":
            name = "Full Payment"
        elif term == "quarterly":
            name = "Quarterly"
        return name

    def action_create_rent_contract_summary_report(self, sheet, record, workbook):
        border_squre = xlwt.Borders()
        border_squre.top = xlwt.Borders.HAIR
        border_squre.left = xlwt.Borders.HAIR
        border_squre.right = xlwt.Borders.HAIR
        border_squre.bottom = xlwt.Borders.HAIR
        border_squre.top_colour = xlwt.Style.colour_map["gray50"]
        border_squre.bottom_colour = xlwt.Style.colour_map["gray50"]
        border_squre.right_colour = xlwt.Style.colour_map["gray50"]
        border_squre.left_colour = xlwt.Style.colour_map["gray50"]
        al = xlwt.Alignment()
        al.horz = xlwt.Alignment.HORZ_CENTER
        al.vert = xlwt.Alignment.VERT_CENTER
        date_format = xlwt.XFStyle()
        date_format.num_format_str = 'mm/dd/yyyy'
        date_format.font.name = "Century Gothic"
        date_format.borders = border_squre
        date_format.alignment = al
        sheet.row(0).height = 1000
        sheet.col(0).width = 6000
        sheet.col(1).width = 7000
        sheet.col(2).width = 6000
        sheet.col(3).width = 6000
        sheet.col(4).width = 6000
        sheet.col(5).width = 6000
        sheet.col(6).width = 6000
        sheet.col(7).width = 6000
        sheet.col(8).width = 6000
        sheet.col(9).width = 6000

        xlwt.add_palette_colour("custom_red", 0x21)
        workbook.set_colour_RGB(0x21, 240, 210, 211)
        xlwt.add_palette_colour("custom_green", 0x23)
        workbook.set_colour_RGB(0x23, 210, 241, 214)

        title = xlwt.easyxf(
            "font: height 440, name Century Gothic, bold on, color_index blue_gray;"
            " align: vert center, horz center;"
            "border: bottom thick, bottom_color sea_green;")
        sub_title = xlwt.easyxf(
            "font: height 185, name Century Gothic, bold on, color_index gray80; "
            "align: vert center, horz center; "
            "border: top hair, bottom hair, left hair, right hair, "
            "top_color gray50, bottom_color gray50, left_color gray50, right_color gray50")
        border_all_right = xlwt.easyxf(
            "align:horz right, vert center;"
            "font:name Century Gothic;"
            "border:  top hair, bottom hair, left hair, right hair, "
            "top_color gray50, bottom_color gray50, left_color gray50, right_color gray50")
        border_all_center = xlwt.easyxf(
            "align:horz center, vert center;"
            "font:name Century Gothic;"
            "border:  top hair, bottom hair, left hair, right hair, "
            "top_color gray50, bottom_color gray50, left_color gray50, right_color gray50")
        running_text = xlwt.easyxf(
            "align:horz center, vert center;"
            "font:name Century Gothic, color_index sea_green, bold on;"
            "border:  top hair, bottom hair, left hair, right hair, "
            "top_color gray50, bottom_color gray50, left_color gray50, right_color gray50")
        cancel_close_text = xlwt.easyxf(
            "align:horz center, vert center;"
            "font:bold on, name Century Gothic, color_index dark_red;"
            "border:  top hair, bottom hair, left hair, right hair, "
            "top_color gray50, bottom_color gray50, left_color gray50, right_color gray50")
        draft_text = xlwt.easyxf(
            "align:horz center, vert center;"
            "font:name Century Gothic, color_index dark_blue, bold on;"
            "border:  top hair, bottom hair, left hair, right hair, "
            "top_color gray50, bottom_color gray50, left_color gray50, right_color gray50")
        total_paid_text = xlwt.easyxf(
            "pattern: pattern solid, fore_colour custom_green;"
            "align:horz right, vert center;"
            "font:name Century Gothic, bold on;"
            "border:  top hair, bottom hair, left hair, right hair, "
            "top_color gray50, bottom_color gray50, left_color gray50, right_color gray50")
        total_remaining_text = xlwt.easyxf(
            "pattern: pattern solid, fore_colour custom_red;"
            "align:horz right, vert center;"
            "font:name Century Gothic, bold on;"
            "border:  top hair, bottom hair, left hair, right hair, "
            "top_color gray50, bottom_color gray50, left_color gray50, right_color gray50")
        expire_text = xlwt.easyxf(
            "align:horz center, vert center;"
            "font:name Century Gothic, color_index olive_ega, bold on;"
            "border:  top hair, bottom hair, left hair, right hair, "
            "top_color gray50, bottom_color gray50, left_color gray50, right_color gray50")


        sheet.write_merge(0, 0, 0, 9, "RENTAL SUMMARY REPORT PER PROPERTIES ", title)
        sheet.write(1, 0, "Properties", sub_title)
        sheet.write(1, 1, "Average Rent/Month", sub_title)
        sheet.write(1, 2, "Total Value Rent/Contracts", sub_title)
        sheet.write(1, 3, "Total Collection", sub_title)
        sheet.write(1, 4, "Remaining Amount", sub_title)
        sheet.write(1, 5, "Overdue Amount", sub_title)


        sheet.write(1, 6, "Total Units", sub_title)
        sheet.write(1, 7, "Total Rented Units", sub_title)
        sheet.write(1, 8, "Total Vacant  Units", sub_title)
        sheet.write(1, 9, "Occupancy Rate", sub_title)

        row = 2
        col = 0

        for rec in record:
            unexpected_amount = 0.0
            scope_of_collection = 0.0
            properties_tenancy = self.env['tenancy.details'].sudo().search(
                [('property_id', '=', rec.id)])

            total_unit_rent = 0.0
            rent_count = 0
            average_rent = 0
            total_collection = 0
            total_contract_rent = 0
            remaining_amount = 0
            overdue_amount = 0
            occupied_unit = 0
            vacant_unit = 0


            for unit in rec.property_unit_ids:
                if unit.sale_lease == "for_tenancy":
                    rent_count += 1
                    total_unit_rent += unit.price
                    if unit.stage == "on_lease":
                        for con in unit.tenancy_ids:
                            if con.contract_type == "running_contract":
                                occupied_unit += 1
                    if unit.stage == "available":
                        vacant_unit += 1
                    for contract in unit.tenancy_ids:
                        if contract.contract_type == "running_contract":
                            total_contract_rent += contract.total_amount
                            total_collection += contract.paid_tenancy
                            remaining_amount += contract.remain_tenancy
                            for i in contract.rent_invoice_ids:
                                if fields.Date.today() > i.invoice_date and i.payment_state == "not_paid":
                                    overdue_amount += i.amount
            if total_unit_rent:
                average_rent = total_unit_rent / rent_count
            occupancy_rate = (occupied_unit /len(rec.property_unit_ids)) * 100 if rent_count else 0

            for i in properties_tenancy:
                if i.contract_type == "running_contract":
                    scope_of_collection += i.remain_tenancy
                if i.contract_type == "close_contract" or i.contract_type == "expire_contract":
                    unexpected_amount += i.total_amount
            sheet.row(row).height = 400
            sheet.write(row, col, rec.name, border_all_center)
            sheet.write(row, col + 1, average_rent, border_all_center)
            sheet.write(row, col + 2, total_contract_rent, border_all_center)
            sheet.write(row, col + 3, total_collection, border_all_center)
            sheet.write(row, col + 4, remaining_amount, border_all_center)
            sheet.write(row, col + 5, overdue_amount, border_all_center)


            sheet.write(row, col + 6, rent_count, border_all_right)
            sheet.write(row, col + 7, occupied_unit, border_all_right)
            sheet.write(row, col + 8, vacant_unit, border_all_right)
            sheet.write(row, col + 9, str(occupancy_rate) + '%', border_all_center)

            row += 1
        sheet.row(row).height = 400

        sheet_contract = workbook.add_sheet(
            'Contract', cell_overwrite_ok=True)

        row = 2
        sheet_contract.row(0).height = 1000
        sheet_contract.col(0).width = 6000
        sheet_contract.col(1).width = 6000
        sheet_contract.col(2).width = 6000
        sheet_contract.col(3).width = 6000
        sheet_contract.col(4).width = 8000
        sheet_contract.col(5).width = 9000

        sheet_contract.write_merge(0, 0, 0, 5, "CONTRACT SUMMARY REPORT PER PROPERTIES ", title)
        sheet_contract.write(1, 0, "Properties", sub_title)
        sheet_contract.write(1, 1, "Total Units", sub_title)
        sheet_contract.write(1, 2, "Running Contract", sub_title)
        sheet_contract.write(1, 3, "Expired Contract ", sub_title)
        sheet_contract.write(1, 4, "Closed Contract ", sub_title)
        sheet_contract.write(1, 5, "Cancelled/ Terminated Contract", sub_title)
        for reco in record:



            col = 0

            total_units = 0
            running_contract_unit = 0
            expired_contract_unit = 0
            closed_contract_unit = 0
            cancelled_contract_unit = 0
            for unit in reco.property_unit_ids:
                if unit.sale_lease == "for_tenancy":
                    total_units += 1
                    for contract in unit.tenancy_ids:
                        if contract.contract_type == "running_contract":
                            running_contract_unit += 1
                        if contract.contract_type == "expire_contract":
                            expired_contract_unit += 1
                        if contract.contract_type == "close_contract":
                            closed_contract_unit += 1
                        if contract.contract_type == "cancel_contract":
                            cancelled_contract_unit += 1

            sheet_contract.write(row, col, reco.name, border_all_center)
            sheet_contract.write(row, col + 1, total_units, border_all_center)
            sheet_contract.write(row, col + 2, running_contract_unit, border_all_center)
            sheet_contract.write(row, col + 3, expired_contract_unit, border_all_center)
            sheet_contract.write(row, col + 4, closed_contract_unit, border_all_center)
            sheet_contract.write(row, col + 5, cancelled_contract_unit, border_all_center)

            row += 1

            sheet_contract.row(row).height = 400

        for reco in record:
            sheet_rec = workbook.add_sheet(
                reco.name, cell_overwrite_ok=True)
            sheet_rec.row(0).height = 1000
            sheet_rec.col(0).width = 6000
            sheet_rec.col(1).width = 6000
            sheet_rec.col(2).width = 6000
            sheet_rec.col(3).width = 6000
            sheet_rec.col(4).width = 6000
            sheet_rec.col(5).width = 6000
            sheet_rec.col(6).width = 6000
            sheet_rec.col(7).width = 6000
            sheet_rec.col(8).width = 6000
            sheet_rec.col(9).width = 6000
            sheet_rec.col(10).width = 6000
            sheet_rec.col(11).width = 6000
            sheet_rec.col(12).width = 6000
            sheet_rec.col(13).width = 6000
            sheet_rec.col(14).width = 6000
            sheet_rec.col(15).width = 6000
            sheet_rec.col(16).width = 6000
            sheet_rec.col(17).width = 6000
            sheet_rec.col(18).width = 6000
            sheet_rec.col(19).width = 6000
            sheet_rec.col(20).width = 6000
            sheet_rec.col(21).width = 6000

            sheet_rec.write_merge(0, 0, 0, 20, "CONTRACT SUMMARY REPORT PER PROPERTIES ", title)
            sheet_rec.write(1, 0, "Reference", sub_title)
            sheet_rec.write(1, 1, "Unit", sub_title)
            sheet_rec.write(1, 2, "Unit Type", sub_title)
            # sheet_rec.write(1, 3, "Property", sub_title)
            sheet_rec.write(1, 3, "Customer", sub_title)
            sheet_rec.write(1, 4, "Landlord", sub_title)
            sheet_rec.write(1, 5, "Broker", sub_title)
            sheet_rec.write(1, 6, "Total Area", sub_title)
            sheet_rec.write(1, 7, "Start Date", sub_title)
            sheet_rec.write(1, 8, "End Date", sub_title)
            sheet_rec.write(1, 9, "Payment Term", sub_title)
            sheet_rec.write(1, 10, "Currency", sub_title)
            sheet_rec.write(1, 11, "Rent", sub_title)
            sheet_rec.write(1, 12, "Security Deposit", sub_title)
            sheet_rec.write(1, 13, "Broker Commission", sub_title)
            sheet_rec.write(1, 14, "Total Bill Amount", sub_title)
            sheet_rec.write(1, 15, "Paid Bill Amount", sub_title)
            sheet_rec.write(1, 16, "Remaining Bill Amount", sub_title)
            sheet_rec.write(1, 17, "Total Amount", sub_title)
            sheet_rec.write(1, 18, "Paid Amount", sub_title)
            sheet_rec.write(1, 19, "Remaining Amount", sub_title)
            sheet_rec.write(1, 20, "Status", sub_title)

            row = 2
            col = 0
            record_details = self.env["property.details"].search([('property_project_id', '=', reco.id)])
            pro_price = 0.0
            pro_total_amount = 0.0
            pro_main_total_payable = 0.0
            pro_main_total_remaining = 0.0
            pro_total_security_deposit = 0.0
            pro_total_broker_commission = 0.0
            pro_total_bill_amount = 0.0
            pro_paid_bill_amount = 0.0
            pro_remaining_bill_amount = 0.0


            for rec in record_details:
                main_total_payable, main_total_remaining, total_security_deposit, total_broker_commission, total_bill_amount, paid_bill_amount, remaining_bill_amount, total_amount = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
                tenency = self.env['tenancy.details'].search([('property_id', '=', rec.id)])
                for i in tenency:
                    total_amount += i.total_amount
                    main_total_payable += i.paid_tenancy
                    main_total_remaining += i.remain_tenancy
                    total_security_deposit += i.deposit_amount
                    total_broker_commission += i.commission
                    total_bill_amount += i.total_bill_amount
                    paid_bill_amount += i.paid_bill_amount
                    remaining_bill_amount += i.remaining_bill_amount
                customer_name = ""
                broker_name = ""
                start_date = ""
                end_date = ""
                term = ""
                contract_type = ""
                if tenency:
                    data = tenency.search([('property_id', '=', rec.id),('contract_type', '=', "running_contract")])
                    if data:
                        customer_name = data[0].tenancy_id.name if data[0].tenancy_id else ""
                        broker_name = data[0].broker_id.name if data[0].broker_id else ""
                        start_date = data[0].start_date if data[0].start_date else ""
                        end_date = data[0].end_date if data[0].end_date else ""
                        term = data[0].payment_term if data[0].payment_term else ""

                    contract_type = tenency[0].contract_type if tenency[0].contract_type else ""


                stage = self.get_rent_stage(contract_type)
                term_data = self.get_payment_term(term)
                pro_price += rec.price
                pro_total_amount += total_amount
                pro_main_total_payable += main_total_payable
                pro_main_total_remaining += main_total_remaining
                pro_total_security_deposit += total_security_deposit
                pro_total_broker_commission += total_broker_commission
                pro_total_bill_amount += total_bill_amount
                pro_paid_bill_amount += paid_bill_amount
                pro_remaining_bill_amount += remaining_bill_amount

                type = self.get_property_type(rec.type)
                unit = self.get_measure_unit(rec.measure_unit)
                # stage = self.get_rent_stage(rec.contract_type)
                # term = self.get_payment_term(rec.payment_term)


                sheet_rec.row(row).height = 400
                sheet_rec.write(row, col, rec.property_seq, border_all_center)
                sheet_rec.write(row, col + 1, rec.name, border_all_center)
                sheet_rec.write(
                    row, col + 2, f"{type} / {rec.property_subtype_id.name}", border_all_center)
                sheet_rec.write(row, col + 3, customer_name, border_all_center)
                sheet_rec.write(row, col + 4, rec.landlord_id.name,
                                border_all_center)#tocheck
                sheet_rec.write(row, col + 5, broker_name,
                                border_all_center)  #tocheck

                sheet_rec.write(
                    row, col + 6, f"{rec.total_area} {unit}", border_all_right)
                sheet_rec.write(row, col + 7, start_date, date_format)
                sheet_rec.write(row, col + 8, end_date, date_format)
                sheet_rec.write(row, col + 9, term_data, border_all_center)
                sheet_rec.write(row, col + 10,
                                f"{self.env.company.currency_id.symbol} ({self.env.company.currency_id.name})",
                                border_all_center)
                sheet_rec.write(row, col + 11, f"{rec.price} {self.env.company.currency_id.symbol} / {rec.rent_unit}",
                                border_all_center)
                sheet_rec.write(
                    row, col + 12, f"{total_security_deposit}", border_all_right)
                sheet_rec.write(
                    row, col + 13, f"{total_broker_commission}", border_all_right)
                sheet_rec.write(
                    row, col + 14, f"{total_bill_amount}", border_all_right)
                sheet_rec.write(
                    row, col + 15, f"{paid_bill_amount}", border_all_right)
                sheet_rec.write(
                    row, col + 16, f"{remaining_bill_amount}", border_all_right)
                sheet_rec.write(
                    row, col + 17, f"{total_amount}", border_all_right)
                sheet_rec.write(
                    row, col + 18, f"{main_total_payable}", border_all_right)
                sheet_rec.write(
                    row, col + 19, f"{main_total_remaining}", border_all_right)
                if contract_type == "new_contract":
                    sheet_rec.write(row, col + 20, stage, draft_text)
                elif contract_type == "running_contract":
                    sheet_rec.write(row, col + 20, stage, running_text)
                elif contract_type in ["cancel_contract", "close_contract"]:
                    sheet_rec.write(row, col + 20, stage, cancel_close_text)
                elif contract_type == "expire_contract":
                    sheet_rec.write(row, col + 20, stage, expire_text)

                row += 1
            sheet_rec.row(row).height = 400
            sheet_rec.write(row, 10, "Totals", sub_title)
            sheet_rec.write(
                row, 11, f"{pro_price}", total_paid_text)
            sheet_rec.write(
                row, 12, f"{pro_total_security_deposit}", total_paid_text)
            sheet_rec.write(
                row, 13, f"{pro_total_broker_commission}", total_paid_text)
            sheet_rec.write(
                row, 14, f"{pro_total_bill_amount}", total_paid_text)
            sheet_rec.write(
                row, 15, f"{pro_paid_bill_amount}", total_paid_text)
            sheet_rec.write(
                row, 16, f"{pro_remaining_bill_amount}", total_remaining_text)
            sheet_rec.write(
                row, 17, f"{pro_total_amount}", total_paid_text)
            sheet_rec.write(
                row, 18, f"{pro_main_total_payable}", total_paid_text)
            sheet_rec.write(
                row, 19, f"{pro_main_total_remaining}", total_remaining_text)


