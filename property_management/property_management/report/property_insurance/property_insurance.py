# Copyright (c) 2022, Nihantra C. Patel and contributors
# For license information, please see license.txt

import frappe
from frappe import _

def execute(filters=None):
	columns, data = get_columns(filters), get_data(filters)
	return columns, data

def get_columns(filters):
	columns = [
				{
					"label": _("Property"),
					"fieldname": "name",
					"fieldtype": "Link",
					"options": "Asset",
					"width": 180
				},
				{
					"label": _("Property Name"),
					"fieldname": "asset_name",
					"fieldtype": "Data",
					"width": 150
				},
				{
					"label": _("Description"),
					"fieldname": "description",
					"fieldtype": "Data",
					"width": 120
				},
				{
					"label": _("Policy Number"),
					"fieldname": "policy_number",
					"fieldtype": "Data",
					"width": 120
				},
				{
					"label": _("Start Date"),
					"fieldname": "start_date",
					"fieldtype": "Date",
					"width": 95
				},
				{
					"label": _("End Date"),
					"fieldname": "end_date",
					"fieldtype": "Date",
					"width": 95
				},
				{
					"label": _("Premium"),
					"fieldname": "premium",
					"fieldtype": "Float",
					"width": 95
				},
				{
					"label": _("Payment Term"),
					"fieldname": "payment_term",
					"fieldtype": "Link",
					"options": "Payment Term",
					"width": 100
				},
				{
					"label": _("Related Company"),
					"fieldname": "related_company",
					"fieldtype": "Data",
					"width": 100
				},
				{
					"label": _("Insurance Company"),
					"fieldname": "insurance_company",
					"fieldtype": "Data",
					"width": 100
				},
				{
					"label": _("Company"),
					"fieldname": "company",
					"fieldtype": "Link",
					"options": "Company",
					"width": 120
				}
			]
	return columns

def get_data(filters):
	return frappe.db.sql(
			"""
			SELECT
				a.company, a.name, a.asset_name, pi.description, pi.policy_number, pi.start_date, pi.end_date, pi.premium, \
				pi.payment_term, pi.related_company, pi.insurance_company, a.company
			FROM
				`tabAsset` as a, `tabProperty Insurance` as pi
			WHERE
					pi.parent = a.name
				AND a.docstatus < 2
				
				{conditions} """.format(
				conditions=get_conditions(filters)
			),
			filters,
			as_dict=1,
		)

def get_conditions(filters):
	conditions = []

	if filters.get("Company"):
		conditions.append(" and a.company = %(company)s")

	if filters.get("name"):
		conditions.append(" and a.name in %(name)s")

	if filters.get("issuance_from_date") and filters.get("issuance_to_date"):
		conditions.append(" and pi.start_date BETWEEN %(issuance_from_date)s AND %(issuance_to_date)s")

	if filters.get("expiry_from_date") and filters.get("expiry_to_date"):
		conditions.append(" and pi.end_date BETWEEN %(expiry_from_date)s AND %(expiry_to_date)s")

	return " ".join(conditions) if conditions else ""