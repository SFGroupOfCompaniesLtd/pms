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
					"label": _("Tenancy ID"),
					"fieldname": "name",
					"fieldtype": "Link",
					"options": "Tenancy",
					"width": 95,
				},
				{
					"label": _("Property"),
					"fieldname": "asset",
					"fieldtype": "Link",
					"options": "Asset",
					"width": 170,
				},
				{
					"label": _("Property Name"),
					"fieldname": "asset_name",
					"fieldtype": "Data",
					"width": 150,
				},
			]
	if filters.get("is_tenant_tenancy") == 1:
		columns.append(
				{
					"label": _("Tenant"),
					"fieldname": "tenant",
					"fieldtype": "Link",
					"options": "Customer",
					"width": 120,
				})
	if filters.get("is_landlord_tenancy") == 1:
		columns.append(
				{
					"label": _("Landlord"),
					"fieldname": "landlord",
					"fieldtype": "Link",
					"options": "Customer",
					"width": 120,
				})
	if filters.get("is_tenant_tenancy") != 1 and filters.get("is_landlord_tenancy") != 1:
		columns.append(
				{
					"label": _("Tenant"),
					"fieldname": "tenant",
					"fieldtype": "Link",
					"options": "Customer",
					"width": 120,
				})
		columns.append(
				{
					"label": _("Landlord"),
					"fieldname": "landlord",
					"fieldtype": "Link",
					"options": "Customer",
					"width": 120,
				})		
	columns.extend(
		[
				{
					"label": _("Start Date"),
					"fieldname": "start_date",
					"fieldtype": "Date",
					"width": 100,
				},
				{
					"label": _("End Date"),
					"fieldname": "end_date",
					"fieldtype": "Date",
					"width": 100,
				},
				{
					"label": _("Rent Type"),
					"fieldname": "rent_type",
					"fieldtype": "Data",
					"width": 95,
				},
				{
					"label": _("Ground Rent"),
					"fieldname": "ground_rent",
					"fieldtype": "Float",
					"width": 120,
				},
				{
					"label": _("Advance Deposit"),
					"fieldname": "advance_deposit",
					"fieldtype": "Float",
					"width": 140,
				},
				{
					"label": _("Invoice"),
					"fieldname": "invoice",
					"fieldtype": "Link",
					"options": "Sales Invoice",
					"width": 130,
				},
				{
					"label": _("Is Paid"),
					"fieldname": "is_paid",
					"fieldtype": "Check",
					"width": 70,
				},
				{
					"label": _("Payment Entry"),
					"fieldname": "payment_entry",
					"fieldtype": "Link",
					"options": "Payment Entry",
					"width": 130,
				},
				{
					"label": _("Company"),
					"fieldname": "company",
					"fieldtype": "Link",
					"options": "Company",
					"width": 130,
				}
			]
		)
	return columns

def get_data(filters):
	return frappe.db.sql(
			"""
			SELECT
				name, asset, asset_name, is_tenant_tenancy, tenant, is_landlord_tenancy, landlord, start_date, end_date, \
				rent_type, ground_rent, advance_deposit, invoice, is_paid, payment_entry, company
			FROM
				`tabTenancy`
			WHERE
				`tabTenancy`.docstatus = 1
				AND start_date BETWEEN %(from_date)s AND %(to_date)s
				{conditions} """.format(
				conditions=get_conditions(filters)
			),
			filters,
			as_dict=1,
		)

def get_conditions(filters):
	conditions = []

	if filters.get("company"):
		conditions.append(" and `tabTenancy`.company = %(company)s")

	if filters.get("name"):
		conditions.append(" and `tabTenancy`.name = %(name)s")

	if filters.get("asset"):
		conditions.append(" and `tabTenancy`.asset = %(asset)s")

	if filters.get("is_tenant_tenancy"):
		conditions.append(" and `tabTenancy`.is_tenant_tenancy = %(is_tenant_tenancy)s")

	if filters.get("tenant"):
		conditions.append(" and `tabTenancy`.tenant in %(tenant)s")

	if filters.get("is_landlord_tenancy"):
		conditions.append(" and `tabTenancy`.is_landlord_tenancy = %(is_landlord_tenancy)s")

	if filters.get("landlord"):
		conditions.append(" and `tabTenancy`.landlord in %(landlord)s")

	if filters.get("rent_type"):
		conditions.append(" and `tabTenancy`.rent_type = %(rent_type)s")

	if filters.get("invoice"):
		conditions.append(" and `tabTenancy`.invoice = %(invoice)s")

	if filters.get("is_paid"):
		conditions.append(" and `tabTenancy`.is_paid = %(is_paid)s")

	if filters.get("payment_entry"):
		conditions.append(" and `tabTenancy`.payment_entry = %(payment_entry)s")


	return " ".join(conditions) if conditions else ""