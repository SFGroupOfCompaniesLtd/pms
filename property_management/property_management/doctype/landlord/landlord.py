# Copyright (c) 2022, Nihantra C. Patel and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document

class Landlord(Document):
	pass

def create_landlord_customer(doc, event):
	landlord_grp = frappe.db.get_list('Customer Group',filters={'name': doc.landlord_group},fields=['name'],as_list=True)

	if not landlord_grp:
		new_landlord_grp = frappe.get_doc(dict(
			doctype = 'Customer Group',
			customer_group_name = doc.landlord_group
		))
		new_landlord_grp.save()


	landlord_dts = frappe.db.get_list('Customer',filters={'landlord': doc.name},fields=['name'],as_list=True)

	if not landlord_dts:
		new_landlord = frappe.get_doc(dict(
			doctype = 'Customer',
			customer_name = doc.name,
			is_landlord= 1,
			is_customer = 0,
			landlord = doc.name,
			customer_group = doc.landlord_group,
			territory = doc.territory
		))
		new_landlord.save()


	upd_landlord = frappe.db.get_value("Customer", {'landlord': doc.name}, 'name')
	
	if upd_landlord:
		ut = frappe.get_doc("Customer",upd_landlord)
		ut.customer_name = doc.name
		ut.landlord = doc.name
		ut.customer_group = doc.landlord_group
		ut.territory = doc.territory
		ut.save()