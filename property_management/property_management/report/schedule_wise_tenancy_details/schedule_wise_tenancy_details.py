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
					"label": _("Schedule Date"),
					"fieldname": "schedule_date",
					"fieldtype": "Date",
					"width": 100,
				},
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
					"label": _("Amount"),
					"fieldname": "amount",
					"fieldtype": "Currency",
					"width": 120,
				},
				{
					"label": _("Pending Amount"),
					"fieldname": "pending_amount",
					"fieldtype": "Currency",
					"width": 120,
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
				t.name, t.asset, t.asset_name, t.tenant, t.landlord, ts.schedule_date, \
				ts.amount, ts.pending_amount, ts.invoice, ts.is_paid, ts.payment_entry, t.company
			FROM
				`tabTenancy` as t, `tabTenant Schedule` as ts
			WHERE
				ts.parent = t.name
			AND t.docstatus = 1
			AND ts.schedule_date BETWEEN %(from_date)s AND %(to_date)s
				{conditions} 
			ORDER BY
				ts.schedule_date asc """.format(
				conditions=get_conditions(filters)
			),
			filters,
			as_dict=1,
		)

def get_conditions(filters):
	conditions = []

	if filters.get("company"):
		conditions.append(" and t.company = %(company)s")

	if filters.get("name"):
		conditions.append(" and t.name = %(name)s")

	if filters.get("asset"):
		conditions.append(" and t.asset = %(asset)s")

	if filters.get("is_tenant_tenancy"):
		conditions.append(" and t.is_tenant_tenancy = %(is_tenant_tenancy)s")

	if filters.get("tenant"):
		conditions.append(" and t.tenant in %(tenant)s")

	if filters.get("is_landlord_tenancy"):
		conditions.append(" and t.is_landlord_tenancy = %(is_landlord_tenancy)s")

	if filters.get("landlord"):
		conditions.append(" and t.landlord in %(landlord)s")

	if filters.get("invoice"):
		conditions.append(" and ts.invoice = %(invoice)s")

	if filters.get("is_paid"):
		conditions.append(" and ts.is_paid = %(is_paid)s")

	if filters.get("payment_entry"):
		conditions.append(" and ts.payment_entry = %(payment_entry)s")


	return " ".join(conditions) if conditions else ""