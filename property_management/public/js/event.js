frappe.ui.form.on("Event", {
	refresh: function(frm) {
		frm.set_query('reference_doctype', "event_participants", function() {
			return {
				"filters": {
					"name": ["in", ["Contact", "Lead", "Customer", "Supplier", "Employee", "Sales Partner", "Asset", "Tenancy"]]
				}
			};
		});
	}
});