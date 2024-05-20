// Copyright (c) 2022, Nihantra C. Patel and contributors
// For license information, please see license.txt
/* eslint-disable */

frappe.query_reports["Property Insurance"] = {
	"filters": [
		{
			"fieldname": "company",
			"label": __("Company"),
			"fieldtype": "Link",
			"options": "Company",
			"default": frappe.defaults.get_user_default("company"),
			"reqd": 1
		},
		{
			"fieldname": "issuance_from_date",
			"label": __("Issuance / Warranty From Date"),
			"fieldtype": "Date"
		},
		{
			"fieldname": "issuance_to_date",
			"label": __("Issuance / Warranty To Date"),
			"fieldtype": "Date"
		},
		{
			"fieldname": "expiry_from_date",
			"label": __("Expiry From Date"),
			"fieldtype": "Date"
		},
		{
			"fieldname": "expiry_to_date",
			"label": __("Expiry To Date"),
			"fieldtype": "Date"
		},
		{
			"fieldname": "name",
			"label": __("Property"),
			"fieldtype": "MultiSelectList",
			"options": "Asset",
			get_data: function(txt) {
				return frappe.db.get_link_options('Asset', txt);
			}
		}
	]
};