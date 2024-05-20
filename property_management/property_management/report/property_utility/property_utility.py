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
					"label": _("Company"),
					"fieldname": "company",
					"fieldtype": "Link",
					"options": "Company",
					"width": 120
				},
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
					"label": _("Utility Type"),
					"fieldname": "utility_type",
					"fieldtype": "Link",
					"options": "Utility Type",
					"width": 120
				},
				{
					"label": _("Contact"),
					"fieldname": "contact",
					"fieldtype": "Link",
					"options": "Contact",
					"width": 120
				},
				{
					"label": _("Reference"),
					"fieldname": "reference",
					"fieldtype": "Data",
					"width": 100
				},
				{
					"label": _("Issuance / Warranty Date"),
					"fieldname": "issuance_date",
					"fieldtype": "Date",
					"width": 115
				},
				{
					"label": _("Expiry Date"),
					"fieldname": "expiry_date",
					"fieldtype": "Date",
					"width": 100
				},
				{
					"label": _("Remarks"),
					"fieldname": "remarks",
					"fieldtype": "Data",
					"width": 200
				}
				
			]
	return columns

def get_data(filters):
	return frappe.db.sql(
			"""
			SELECT
				a.company, a.name, a.asset_name, u.utility_type, u.contact, u.reference, u.issuance_date, u.expiry_date, u.remarks
			FROM
				`tabAsset` as a, `tabUtility` as u
			WHERE
					u.parent = a.name
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

	if filters.get("utility_type"):
		conditions.append(" and u.utility_type in %(utility_type)s")

	if filters.get("issuance_from_date") and filters.get("issuance_to_date"):
		conditions.append(" and u.issuance_date BETWEEN %(issuance_from_date)s AND %(issuance_to_date)s")

	if filters.get("expiry_from_date") and filters.get("expiry_to_date"):
		conditions.append(" and u.expiry_date BETWEEN %(expiry_from_date)s AND %(expiry_to_date)s")

	return " ".join(conditions) if conditions else ""