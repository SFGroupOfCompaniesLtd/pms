# Copyright (c) 2022, Nihantra C. Patel and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document

class PropertyManager(Document):
	pass

def create_property_manager_customer(doc, event):
	pro_mngr_grp = frappe.db.get_list('Customer Group',filters={'name': doc.property_manager_group},fields=['name'],as_list=True)

	if not pro_mngr_grp:
		new_pro_mngr_grp = frappe.get_doc(dict(
			doctype = 'Customer Group',
			customer_group_name = doc.property_manager_group
		))
		new_pro_mngr_grp.save()


	pro_mngr = frappe.db.get_list('Customer',filters={'property_manager': doc.name},fields=['name'],as_list=True)

	if not pro_mngr:
		new_pro_mngr = frappe.get_doc(dict(
			doctype = 'Customer',
			customer_name = doc.name,
			is_property_manager = 1,
			is_customer = 0,
			property_manager = doc.name,
			customer_group = doc.property_manager_group,
			territory = doc.territory
		))
		new_pro_mngr.save()


	upd_pro_mngr = frappe.db.get_value("Customer", {'property_manager': doc.name}, 'name')
	
	if upd_pro_mngr:
		upm = frappe.get_doc("Customer",upd_pro_mngr)
		upm.customer_name = doc.name
		upm.property_manager = doc.name
		upm.customer_group = doc.property_manager_group
		upm.territory = doc.territory
		upm.save()