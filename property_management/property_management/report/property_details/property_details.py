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
					"width": 100
				},
				{
					"label": _("Property"),
					"fieldname": "name",
					"fieldtype": "Link",
					"options": "Asset",
					"width": 170
				},
				{
					"label": _("Start Date"),
					"fieldname": "start_date",
					"fieldtype": "Date",
					"width": 100
				},
				{
					"label": _("Property Name"),
					"fieldname": "asset_name",
					"fieldtype": "Data",
					"width": 150
				},
				{
					"label": _("Property Category"),
					"fieldname": "asset_category",
					"fieldtype": "Link",
					"options": "Asset Category",
					"width": 150
				},
				{
					"label": _("Property Status"),
					"fieldname": "property_status",
					"fieldtype": "Data",
					"width": 110
				},
				{
					"label": _("Furnishing"),
					"fieldname": "furnishing",
					"fieldtype": "Data",
					"width": 110
				},
				{
					"label": _("Bedrooms"),
					"fieldname": "bedrooms",
					"fieldtype": "Data",
					"width": 90
				},
				{
					"label": _("Bathrooms"),
					"fieldname": "bathrooms",
					"fieldtype": "Data",
					"width": 92
				},
				{
					"label": _("Parking"),
					"fieldname": "parking",
					"fieldtype": "Data",
					"width": 72
				},
				{
					"label": _("Facing"),
					"fieldname": "facing",
					"fieldtype": "Data",
					"width": 68
				},
				{
					"label": _("Total Towers"),
					"fieldname": "no_of_towers",
					"fieldtype": "Int",
					"width": 105
				},
				{
					"label": _("Location"),
					"fieldname": "location",
					"fieldtype": "Link",
					"options": "Location",
					"width": 100
				},
				{
					"label": _("Rent Type"),
					"fieldname": "rent_type",
					"fieldtype": "Data",
					"width": 90
				},
				{
					"label": _("Ground Rent"),
					"fieldname": "ground_rent",
					"fieldtype": "Float",
					"width": 120
				},
				{
					"label": _("GFA sqft"),
					"fieldname": "gfa_sqft",
					"fieldtype": "Float",
					"width": 80
				},
				{
					"label": _("GFA m"),
					"fieldname": "gfa_m",
					"fieldtype": "Float",
					"width": 80
				},
				{
					"label": _("Unit Price"),
					"fieldname": "unit_price",
					"fieldtype": "Float",
					"width": 90
				},
				{
					"label": _("Total Price"),
					"fieldname": "total_price",
					"fieldtype": "Float",
					"width": 90
				},
				{
					"label": _("Property Manager"),
					"fieldname": "property_manager",
					"fieldtype": "Link",
					"options": "Customer",
					"width": 100
				},
			]
	return columns

def get_data(filters):
	return frappe.db.sql(
			"""
			SELECT
				company, name, asset_name, start_date, asset_category, \
				CASE 
					WHEN property_status = "Available" THEN '<c style="color:#2f9d58;">' "Available" '</c>'
					WHEN property_status = "Booked" THEN '<c style="color:#1579d0;">' "Booked" '</c>'
					WHEN property_status = "Rent/Lease" THEN '<c style="color:#f8814f;">' "Rent/Lease" '</c>'
					WHEN property_status = "Closed" THEN '<c style="color:#e24c4c;">' "Closed" '</c>'
				END as property_status,
				furnishing, bedrooms, bathrooms, parking, facing,no_of_towers, \
				location, rent_type, ground_rent, gfa_sqft, gfa_m, unit_price, total_price, property_manager
			FROM
				`tabAsset`
			WHERE
				`tabAsset`.docstatus < 2
				AND start_date BETWEEN %(from_date)s AND %(to_date)s
				{conditions} """.format(
				conditions=get_conditions(filters)
			),
			filters,
			as_dict=1,
		)

def get_conditions(filters):
	conditions = []

	if filters.get("Company"):
		conditions.append(" and `tabAsset`.company = %(company)s")

	if filters.get("name"):
		conditions.append(" and `tabAsset`.name in %(name)s")

	if filters.get("location"):
		conditions.append(" and `tabAsset`.location in %(location)s")

	if filters.get("property_manager"):
		conditions.append(" and `tabAsset`.property_manager in %(property_manager)s")

	if filters.get("bedrooms"):
		conditions.append(" and `tabAsset`.bedrooms = %(bedrooms)s")

	if filters.get("furnishing"):
		conditions.append(" and `tabAsset`.furnishing = %(furnishing)s")

	if filters.get("property_status"):
		conditions.append(" and `tabAsset`.property_status = %(property_status)s")

	if filters.get("is_existing_asset"):
		conditions.append(" and `tabAsset`.is_existing_asset = %(is_existing_asset)s")


	return " ".join(conditions) if conditions else ""