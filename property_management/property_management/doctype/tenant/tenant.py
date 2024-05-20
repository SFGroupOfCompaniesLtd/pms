# Copyright (c) 2022, Nihantra C. Patel and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document

class Tenant(Document):
	pass

def create_tenant_customer(doc, event):
	tenant_grp = frappe.db.get_list('Customer Group',filters={'name': doc.tenant_group},fields=['name'],as_list=True)

	if not tenant_grp:
		new_tenant_grp = frappe.get_doc(dict(
			doctype = 'Customer Group',
			customer_group_name = doc.tenant_group
		))
		new_tenant_grp.save()


	tenant_dts = frappe.db.get_list('Customer',filters={'tenant': doc.name},fields=['name'],as_list=True)

	if not tenant_dts:
		new_tenant = frappe.get_doc(dict(
			doctype = 'Customer',
			customer_name = doc.name,
			is_tenant= 1,
			is_customer = 0,
			tenant = doc.name,
			customer_group = doc.tenant_group,
			territory = doc.territory
		))
		new_tenant.save()


	upd_tenant = frappe.db.get_value("Customer", {'tenant': doc.name}, 'name')
	
	if upd_tenant:
		ut = frappe.get_doc("Customer",upd_tenant)
		ut.customer_name = doc.name
		ut.tenant = doc.name
		ut.customer_group = doc.tenant_group
		ut.territory = doc.territory
		ut.save()