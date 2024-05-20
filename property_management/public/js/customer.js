frappe.ui.form.on('Customer', {
	refresh(frm) {
		frm.add_custom_button(__("Tenancy Details"), function() {
		    if (frm.doc.is_tenant == 1) {
		        frappe.route_options = {
		            "is_tenant_tenancy": 1,
    				"tenant": frm.doc.name
    			};
		    }
		    else if (frm.doc.is_landlord == 1) {
		        frappe.route_options = {
		            "is_landlord_tenancy": 1,
    				"landlord": frm.doc.name
    			};
		    }
		    frappe.set_route("query-report", "Tenancy Details");
		}, __("View"));
		
		frm.add_custom_button(__("Schedule-wise Tenancy Details"), function() {
		    if (frm.doc.is_tenant == 1) {
		        frappe.route_options = {
		            "is_tenant_tenancy": 1,
    				"tenant": frm.doc.name
    			};
		    }
		    else if (frm.doc.is_landlord == 1) {
		        frappe.route_options = {
		            "is_landlord_tenancy": 1,
    				"landlord": frm.doc.name
    			};
		    }
		    frappe.set_route("query-report", "Schedule-wise Tenancy Details");
		}, __("View"));
	}
});
