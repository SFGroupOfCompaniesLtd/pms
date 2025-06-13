# -*- coding: utf-8 -*-
# Copyright 2020-Today TechKhedut.
# Part of TechKhedut. See LICENSE file for full copyright and licensing details.
from odoo import models, fields, api


class UserTypes(models.Model):
    _inherit = 'res.partner'

    user_type = fields.Selection([('landlord', 'LandLord'),
                                  ('customer', 'Customer'),
                                  ('broker', 'Broker')],
                                 string='User Type')
    properties_count = fields.Integer(string='Properties Count', compute='_compute_properties_count')
    defaulter_count = fields.Integer(string='Defaulter Count', compute='_compute_defaulter_count')
    properties_ids = fields.One2many('property.details', 'landlord_id', string='Properties')
    brokerage_company_id = fields.Many2one('res.company', string=' Brokerage Company',
                                           default=lambda self: self.env.company)
    currency_id = fields.Many2one('res.currency', related='brokerage_company_id.currency_id',
                                  string='Currency')

    # Customer Fields
    is_tenancy = fields.Boolean(string='Property Renting')
    is_sold_customer = fields.Boolean(string='Property Buyer')

    # Broker Fields
    tenancy_ids = fields.One2many('tenancy.details', 'broker_id', string='Tenancy ')
    property_sold_ids = fields.One2many('property.vendor', 'broker_id', string="Sold Commission")
    defaulter = fields.Boolean('Defaulter')

    @api.depends('properties_ids')
    def _compute_properties_count(self):
        for rec in self:
            count = self.env['property.details'].search_count([('landlord_id', 'in', [rec.id])])
            rec.properties_count = count

    def _compute_defaulter_count(self):
        for rec in self:
            count = self.env['account.move'].search_count([('partner_id', '=', rec.id), ('defaulter_move', '=', True)])
            rec.defaulter_count = count

    def action_properties(self):
        return {
            'type': 'ir.actions.act_window',
            'name': 'Properties',
            'res_model': 'property.details',
            'domain': [('landlord_id', '=', self.id)],
            'context': {'default_landlord_id': self.id},
            'view_mode': 'list,form',
            'target': 'current'
        }

    def action_view_defaulter_invoices(self):
        self.ensure_one()
        action = self.env["ir.actions.actions"]._for_xml_id("account.action_move_out_invoice_type")
        all_child = self.with_context(active_test=False).search([('id', 'child_of', self.ids)])
        action['domain'] = [
            ('move_type', 'in', ('out_invoice', 'out_refund')),
            ('partner_id', 'in', all_child.ids), ('defaulter_move', '=', True)
        ]
        action['context'] = {'default_move_type': 'out_invoice', 'move_type': 'out_invoice', 'journal_type': 'sale',
                             'search_default_unpaid': 1}
        return action
