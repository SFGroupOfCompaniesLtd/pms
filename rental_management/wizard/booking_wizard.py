from odoo import fields, api, models
import requests
import json
from dateutil.relativedelta import relativedelta

import logging as logger
_logger = logger.getLogger(__name__)

class BookingWizard(models.TransientModel):
    _name = 'booking.wizard'
    _description = 'Create Booking While Property on Sale'

    customer_id = fields.Many2one('res.partner', string='Customer', domain="[('user_type','=','customer')]")
    property_id = fields.Many2one('property.details', string='Property')
    price = fields.Monetary(related="property_id.price")
    company_id = fields.Many2one('res.company', string='Company', default=lambda self: self.env.company)
    currency_id = fields.Many2one('res.currency', related='company_id.currency_id', string='Currency')
    book_price = fields.Monetary(string="Advance")
    ask_price = fields.Monetary(string="Customer Price")
    is_any_broker = fields.Boolean(string='Any Broker?')
    broker_id = fields.Many2one('res.partner', string='Broker', domain=[('user_type', '=', 'broker')])
    commission_type = fields.Selection([('f', 'Fix'), ('p', 'Percentage')], string="Commission Type")
    broker_commission = fields.Monetary(string='Commission')
    broker_commission_percentage = fields.Float(string='Percentage')
    commission_from = fields.Selection([('customer', 'Customer'),
                                        ('landlord', 'Landlord',)],
                                       default='customer', string="Commission From")
    from_inquiry = fields.Boolean('From Enquiry')
    note = fields.Text(string="Note", translate=True)
    lead_id = fields.Many2one('crm.lead', string="Enquiry", domain="[('property_id','=',property_id)]")

    # Maintenance and utility Service
    is_any_maintenance = fields.Boolean(related="property_id.is_maintenance_service")
    total_maintenance = fields.Monetary(related="property_id.total_maintenance")
    is_utility_service = fields.Boolean(related="property_id.is_extra_service")
    total_service = fields.Monetary(related="property_id.extra_service_cost")

    # Booking Item
    booking_item_id = fields.Many2one('product.product', string="Booking Item")
    broker_item_id = fields.Many2one('product.product', string="Broker Item")

    @api.model
    def default_get(self, fields):
        res = super(BookingWizard, self).default_get(fields)
        active_id = self._context.get('active_id')
        property_id = self.env['property.details'].browse(active_id)
        default_broker_item = self.env['ir.config_parameter'].sudo().get_param(
            'rental_management.account_broker_item_id')
        default_deposit_item = self.env['ir.config_parameter'].sudo().get_param(
            'rental_management.account_deposit_item_id')
        res['property_id'] = property_id.id
        res['ask_price'] = property_id.price
        res['booking_item_id'] = int(default_deposit_item) if default_deposit_item else self.env.ref(
            'rental_management.property_product_2').id
        res['broker_item_id'] = int(default_broker_item) if default_broker_item else self.env.ref(
            'rental_management.property_product_3').id
        return res

    def create_booking_action(self):
        invoice_post_type = self.env['ir.config_parameter'].sudo().get_param('rental_management.invoice_post_type')
        booking_due_days = self.env['ir.config_parameter'].sudo().get_param(
            'rental_management.booking_due_days')
        gva_url = self.env['ir.config_parameter'].sudo().get_param('rental_management.gva_url')
        gva_access_token = self.env['ir.config_parameter'].sudo().get_param('rental_management.gva_access_token')
        malipopay = self.env['ir.config_parameter'].sudo().get_param('rental_management.malipopay')
        self.customer_id.user_type = "customer"
        lead = self._context.get('from_crm')
        book_invoice_id = self.env['account.move']
        data = {
            'customer_id': self.customer_id.id,
            'property_id': self.property_id.id,
            'book_price': self.book_price * (-1),
            'ask_price': self.ask_price,
            'is_any_broker': self.is_any_broker,
            'broker_id': self.broker_id.id,
            'commission_type': self.commission_type,
            'broker_commission': self.broker_commission,
            'broker_commission_percentage': self.broker_commission_percentage,
            'stage': 'booked',
            'commission_from': self.commission_from,
            'booking_item_id': self.booking_item_id.id,
            'broker_item_id': self.broker_item_id.id,
        }
        booking_id = self.env['property.vendor'].create(data)
        # self.property_id.sold_booking_id = booking_id.id
        mail_template = self.env.ref(
            'rental_management.property_book_mail_template', raise_if_not_found=False)
        if mail_template:
            mail_template.send_mail(booking_id.id, force_send=True)
        if not booking_id.book_price == 0:
            record = {
                'product_id': self.booking_item_id.id,
                'name': 'Booked Amount of   ' + booking_id.property_id.name,
                'quantity': 1,
                'price_unit': self.book_price
            }
            invoice_lines = [(0, 0, record)]
            today_date = fields.date.today()
            data = {
                'partner_id': booking_id.customer_id.id,
                'move_type': 'out_invoice',
                'booking_invoice':True,
                'invoice_date': today_date,
                'invoice_date_due': today_date + relativedelta(days=int(booking_due_days)) ,
                'invoice_line_ids': invoice_lines
            }
            book_invoice_id = self.env['account.move'].sudo().create(data)
            book_invoice_id.sold_id = booking_id.id
            if invoice_post_type == 'automatically':
                book_invoice_id.action_post()
            booking_id.book_invoice_id = book_invoice_id.id
            booking_id.book_invoice_state = True
        booking_id.property_id.stage = 'booked'
        booking_id.stage = 'booked'
        if book_invoice_id and malipopay:
            if book_invoice_id.state == 'posted' and self.property_id.type == 'land' and self.property_id.sale_lease == 'for_sale':
                data = {
                    "reference": book_invoice_id.name,
                    "description": 'Booking invoice of '+ self.property_id.name,
                    "channel": "Bank Transfer",
                    "amount": book_invoice_id.amount_total,
                    "sold_invoice_id": book_invoice_id.id,
                    "customer": {
                        "name": self.customer_id.name,
                        "email": self.customer_id.email,
                        "phoneNumber": self.customer_id.phone.replace(' ','').replace('+','') or self.customer_id.mobile.replace(' ','').replace('+','')
                    },
                    "land_details": {
                        "plot_number": self.property_id.name,
                        "location": self.property_id.region_id.name if self.property_id.region_id else '',
                        "booking_amount": book_invoice_id.amount_total,
                        "currency": "USD"
                    }
                }
                _logger.info('Malipopay Push Data\n\n %s', data)
                _logger.info('\n\nGVA ACCESS %s', gva_access_token)
                if data:
                    headers = {"Authorization": gva_access_token,
                               'Content-type': 'application/json'}

                    user_response_data = requests.post(
                        gva_url, headers=headers, data=json.dumps(data))

                    return_data = user_response_data.json()
                    _logger.info(user_response_data,'%s RESPONSE')
                    _logger.info(return_data,'%s WHATS THE RETURN')
                    if return_data.get('data'):
                        gva_reference = return_data['data'].get('reference')
                        gva_payment_id = return_data['data'].get('payment_id')
                        booking_id.gva_reference = gva_reference
                        booking_id.gva_payment_id = gva_payment_id
                        booking_id.book_invoice_id.control_number = gva_payment_id
        return {
            'type': 'ir.actions.act_window',
            'name': 'Property Booking',
            'res_model': 'property.vendor',
            'res_id': booking_id.id,
            'view_mode': 'form,list',
            'target': 'current'
        }

    @api.onchange('from_inquiry')
    def _onchange_property_sale_inquiry(self):
        inquiry_ids = self.env['sale.inquiry'].search(
            [('property_id', '=', self.property_id.id)]).mapped('id')
        for rec in self:
            if not rec.from_inquiry:
                return
            return {'domain': {'inquiry_id': [('id', 'in', inquiry_ids)]}}

    @api.onchange('lead_id')
    def _onchange_ask_price(self):
        for rec in self:
            if not rec.from_inquiry and not rec.lead_id:
                return
            rec.ask_price = rec.lead_id.ask_price
            rec.note = rec.lead_id.description
            rec.customer_id = rec.lead_id.partner_id.id
