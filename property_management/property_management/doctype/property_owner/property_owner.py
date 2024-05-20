# Copyright (c) 2022, Nihantra C. Patel and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document

class PropertyOwner(Document):
	pass

def create_property_owner_customer(doc, event):
	pro_owner_grp = frappe.db.get_list('Customer Group',filters={'name': doc.property_owner_group},fields=['name'],as_list=True)

	if not pro_owner_grp:
		new_pro_owner_grp = frappe.get_doc(dict(
			doctype = 'Customer Group',
			customer_group_name = doc.property_owner_group
		))
		new_pro_owner_grp.save()


	pro_owner = frappe.db.get_list('Customer',filters={'property_owner': doc.name},fields=['name'],as_list=True)

	if not pro_owner:
		new_pro_owner = frappe.get_doc(dict(
			doctype = 'Customer',
			customer_name = doc.name,
			is_property_owner = 1,
			is_customer = 0,
			property_owner = doc.name,
			customer_group = doc.property_owner_group,
			territory = doc.territory
		))
		new_pro_owner.save()


	upd_pro_owner = frappe.db.get_value("Customer", {'property_owner': doc.name}, 'name')
	
	if upd_pro_owner:
		upo = frappe.get_doc("Customer",upd_pro_owner)
		upo.customer_name = doc.name
		upo.property_owner = doc.name
		upo.customer_group = doc.property_owner_group
		upo.territory = doc.territory
		upo.save()