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
					"width": 200
				},
				{
					"label": _("Property Name"),
					"fieldname": "asset_name",
					"fieldtype": "Data",
					"width": 200
				},
				{
					"label": _("Place Name"),
					"fieldname": "place_name",
					"fieldtype": "Data",
					"width": 200
				},
				{
					"label": _("Place Type"),
					"fieldname": "place_type",
					"fieldtype": "Link",
					"options": "Place Type",
					"width": 150
				},
				{
					"label": _("Distance (km)"),
					"fieldname": "distance",
					"fieldtype": "Float",
					"width": 120
				},
			]
	return columns

def get_data(filters):
	return frappe.db.sql(
			"""
			SELECT
				a.company, a.name, a.asset_name, np.place_name, np.place_type, np.distance
			FROM
				`tabAsset` as a, `tabNearest Places` as np
			WHERE
					np.parent = a.name
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

	if filters.get("place_type"):
		conditions.append(" and np.place_type in %(place_type)s")
	
	return " ".join(conditions) if conditions else ""