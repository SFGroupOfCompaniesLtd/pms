# Copyright (c) 2022, Nihantra C. Patel and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
from frappe.utils import flt

class Tenancy(Document):
	pass

def tenant_schedule(doc, event):
    for d in doc.get('tenant_schedule'):
        if d.schedule_date:
            event = frappe.db.get_list('Event',filters={'tenant_schedule_id': d.tenant_schedule_id},fields=['tenant_schedule_id'],as_list=True)
            if not event:
                new_event = frappe.get_doc(dict(
                    doctype = 'Event',
                    starts_on = d.schedule_date,
                    subject = doc.asset_name +' - '+ d.schedule_date,
                    asset_id = doc.asset,
                    tenant_schedule_id = d.tenant_schedule_id
                ))
                new_event.append('event_participants', {
                    'reference_doctype': "Asset", 'reference_docname': doc.asset
                    })
                new_event.append('event_participants', {
                    'reference_doctype': "Tenancy", 'reference_docname': doc.name
                    })
                new_event.save()

@frappe.whitelist()
def create_invoice(tenant, prt, prt_name, amt, custom_tenancy_id):
    items = [
        {"item_code": "Sujlam", "qty": 1, "rate": 100},
        
    ]

    ass_item = frappe.db.get_value('Item', {'asset': prt}, ['item_code'])
    doc = frappe.new_doc("Sales Invoice")
    doc.customer = tenant  
    doc.property = prt  
    doc.property_name = prt_name
    doc.custom_tenancy_id = custom_tenancy_id
    for item in items:
        doc.append("items", {
            "item_code": ass_item,
            "qty": 1,
            "rate": amt,
            "asset": prt
        })
    #doc.insert()
    doc.insert(ignore_permissions=True)
    doc.submit()
    frappe.db.commit()
    return doc.name  

@frappe.whitelist()
def create_paymententry(doc, invoice_name, party, payment_amount, paid_amount,received_amount, paid_to, schedule_date,invoice_ref):
    if invoice_name:
            payment_entry = frappe.get_doc(dict(
                doctype = 'Payment Entry',
                payment_type = 'Receive',
                party_type = 'Customer',
                party = party,
                posting_date = frappe.utils.today(),
                custom_schedule_date = schedule_date,
                custom_invoice_ref = invoice_ref,
                payment_amount = abs(float(payment_amount)),
                paid_amount = abs(float(payment_amount)),
                received_amount = received_amount,
                paid_to = paid_to
            ))
            payment_entry.insert()
            payment_entry.submit()
            frappe.db.commit()
    return payment_entry.name