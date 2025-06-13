# -*- coding: utf-8 -*-
# Copyright 2020-Today TechKhedut.
# Part of TechKhedut. See LICENSE file for full copyright and licensing details.
import base64
from odoo import api, fields, models, tools, _
from odoo.tools.image import is_image_size_above
from odoo.exceptions import ValidationError
from odoo.addons.web_editor.tools import get_video_embed_code, get_video_thumbnail


class PayHistory(models.Model):
    _name = 'pay.history'
    _description = 'Property Sale Payment Details'
    _inherit = ['mail.thread', 'mail.activity.mixin']

    move_id = fields.Many2one('account.move',string='Invoice')
    payment_date = fields.Date(string='Payment Date')
    amount = fields.Float(string='Amount')
    register_payment = fields.Boolean(string='Payment Registered')

    def register_pay(self):
        pay_histories = self.env['pay.history'].search([('register_payment','=',False)])
        for rec in pay_histories:
            wizard = self.env['account.payment.register'].with_context(active_model='account.move',
                                                                       active_ids=rec.move_id.ids).sudo().create(
                {'payment_date': rec.payment_date, 'amount': rec.amount, 'company_id': rec.move_id.company_id.id})
            wizard.sudo().action_create_payments()
            rec.register_payment = True
        