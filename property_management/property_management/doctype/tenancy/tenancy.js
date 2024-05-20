// Copyright (c) 2022, Nihantra C. Patel and contributors
// For license information, please see license.txt

frappe.ui.form.on("Tenancy", {
    refresh: function(frm, cdt, cdn) {
    	// if (frm.doc.docstatus != 1) {
	        frm.add_custom_button(__("Get Schedule"), function() {
	            frm.clear_table("tenant_schedule");
	            frappe.call({
	                method: "frappe.client.get",
	                args: {
	                    doctype: "Asset Depreciation Schedule",
	                filters: {
	                	"asset": cur_frm.doc.asset
	                },
	                fieldname:["asset"]
	                },	               
	                callback(r) {
	                    if (r.message) {
	                        for (var row in r.message.depreciation_schedule) {
	                            var child = frm.add_child("tenant_schedule");
	                            var rms = r.message.depreciation_schedule[row];
	                            frappe.model.set_value(child.doctype, child.name, "schedule_date", rms.schedule_date);
	                            frappe.model.set_value(child.doctype, child.name, "tenant_schedule_id", rms.name);
	                            frappe.model.set_value(child.doctype, child.name, "amount", rms.depreciation_amount);
	                            frappe.model.set_value(child.doctype, child.name, "total_amount", rms.accumulated_depreciation_amount);
	                            frm.refresh_field("tenant_schedule");
	                        }
	                    }
	                }
	            });
	        });
		// }
    }
});

frappe.ui.form.on('Tenancy', {
	refresh: function(frm) {
		frm.fields_dict.tenant_schedule.grid.wrapper.find('.grid-remove-rows').hide();
		frm.fields_dict.tenant_schedule.grid.wrapper.find('.grid-add-row').hide();
	}
});

frappe.ui.form.on('Tenant Schedule', {
	form_render(frm, cdt, cdn) {
        frm.fields_dict.tenant_schedule.grid.wrapper.find('.grid-delete-row').hide();
        frm.fields_dict.tenant_schedule.grid.wrapper.find('.grid-duplicate-row').hide();
        frm.fields_dict.tenant_schedule.grid.wrapper.find('.grid-move-row').hide();
        frm.fields_dict.tenant_schedule.grid.wrapper.find('.grid-append-row').hide();
        frm.fields_dict.tenant_schedule.grid.wrapper.find('.grid-insert-row-below').hide();
        frm.fields_dict.tenant_schedule.grid.wrapper.find('.grid-insert-row').hide();
    }
});

frappe.ui.form.on('Tenant Schedule', {
	invoice: function(frm, cdt, cdn) {
		var d = locals[cdt][cdn];
		frappe.model.set_value(cdt,cdn,'pending_amount', d.amount - d.invoice_amount);
	},
	is_invoice: function(frm,cdt,cdn) {
		var e = locals[cdt][cdn];
		if(e.is_invoice == 1){
			createInvoice(frm,cdt,cdn);
		}
	}
});

frappe.ui.form.on('Tenant Schedule', {
	is_paid: function(frm,cdt,cdn) {
		var e = locals[cdt][cdn];
		if(e.is_paid == 1){
			create_paymententry(frm,cdt,cdn);
		}
	}
});

function createInvoice(frm,cdt,cdn) {
	var row = locals[cdt][cdn];
    frappe.call({
        method: 'property_management.property_management.doctype.tenancy.tenancy.create_invoice',
        args: {
            tenant: frm.doc.tenant,
            prt: frm.doc.asset,
            prt_name: frm.doc.asset_name,
            amt: row.amount,
            custom_tenancy_id: frm.doc.name
        },
        callback: function(response) {
            if (response.message) {
                frappe.msgprint(`${response.message} Invoice created successfully!`);
                frappe.model.set_value(cdt,cdn,'invoice', response.message);
                frm.save('Update');
            } else {
                frappe.msgprint('Failed to create invoice');
            }
        }
    });
}


function create_paymententry(frm,cdt,cdn) {
	var row = locals[cdt][cdn];
	invoice_name = row.invoice
	payment_amount = row.invoice_amount
	schedule_date = row.schedule_date
    frappe.call({
        method: 'property_management.property_management.doctype.tenancy.tenancy.create_paymententry',
        args: {
            party: frm.doc.tenant,
            payment_amount: row.invoice_amount,
            paid_amount: row.invoice_amount,
            received_amount: row.invoice_amount,
            paid_to: 'Cash - S',
            invoice_name: row.invoice,
            doc: frm.doc.name,
            schedule_date: schedule_date,
            invoice_ref: row.invoice
        },
        callback: function(response) {
            if (response.message) {
                frappe.msgprint(`${response.message} Payment Entry created successfully!`);
                frappe.model.set_value(cdt,cdn,'payment_entry', response.message);
                frm.save('Update');
            } else {
                frappe.msgprint('Failed to create payment entry');
            }
        }
    });
}

frappe.ui.form.on('Tenancy', {
	is_tenant_tenancy: function(frm) {
		if (frm.doc.is_tenant_tenancy == 1) {
			frm.set_value('naming_series', "TT.-");
			frm.set_value('is_landlord_tenancy', 0);
			frm.set_value('landlord', undefined);
		}
	},
	is_landlord_tenancy: function(frm) {
		if (frm.doc.is_landlord_tenancy == 1) {
			frm.set_value('naming_series', "LT.-");
			frm.set_value('is_tenant_tenancy', 0);
			frm.set_value('tenant', undefined);
		}
	}
});

frappe.ui.form.on('Tenancy', {
	validate: function(frm) {
		if (frm.doc.is_tenant_tenancy == 0 && frm.doc.is_landlord_tenancy == 0) {
			frappe.msgprint({
			    title: __('Error'),
			    indicator: 'red',
			    message: __('Please select one of the <b>Is Tenant Tenancy</b> or <b>Is Landlord Tenancy</b>')
			});
			validated = false;
		}
	}
});

frappe.ui.form.on("Tenancy", "refresh", function(frm) {
	frm.set_query("asset", function() {
		return {
			"filters": {
				"docstatus": 1
			}
        };
    }),
	frm.set_query("tenant", function() {
		return {
			"filters": {
				"is_tenant": 1
			}
        };
    }),
    frm.set_query("landlord", function() {
		return {
			"filters": {
				"is_landlord": 1
			}
        };
    });
});