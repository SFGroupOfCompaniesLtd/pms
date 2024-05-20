import frappe

def rent_item(doc, event):
    rent_group = frappe.db.get_list('Item Group',filters={'name': "Rent"},fields=['name'],as_list=True)
    if not rent_group:
        new_rent_group = frappe.get_doc(dict(
            doctype = 'Item Group',
            item_group_name = "Rent"
        ))
        new_rent_group.save()

    rent = frappe.db.get_list('Item',filters={'asset': doc.name},fields=['asset'],as_list=True)
    if not rent:
        new_rent_item = frappe.get_doc(dict(
            doctype = 'Item',
            item_code = doc.name + ': ' + doc.asset_name + ' - ' + "Rent",
            item_name = doc.name + ': ' + doc.asset_name + ' - ' + "Rent",
            item_group = "Rent",
            stock_uom = "Nos",
            is_stock_item = 0,
            include_item_in_manufacturing = 0,
            is_fixed_asset = 0,
            is_rent_item = 1,
            asset = doc.name
        ))
        new_rent_item.save()
