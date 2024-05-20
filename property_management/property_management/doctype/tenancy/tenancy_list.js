frappe.listview_settings['Tenancy'] = {
    onload: function(listview) {
       listview.page.fields_dict.asset.get_query = function() {
            return {
				"filters": {
					"docstatus": 1
				}
			};
		};
        listview.page.fields_dict.tenant.get_query = function() {
            return {
				"filters": {
					"is_tenant": 1
				}
			};
		};
		listview.page.fields_dict.landlord.get_query = function() {
			return {
				"filters": {
					"is_landlord": 1
				}
			};
		};		
	}
};