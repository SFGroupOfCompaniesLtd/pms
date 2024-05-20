// Copyright (c) 2022, Nihantra C. Patel and contributors
// For license information, please see license.txt
/* eslint-disable */

frappe.query_reports["Nearest Places of Property"] = {
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
			"fieldname": "name",
			"label": __("Property"),
			"fieldtype": "MultiSelectList",
			"options": "Asset",
			get_data: function(txt) {
				return frappe.db.get_link_options('Asset', txt);
			}
		},
		{
			"fieldname": "place_type",
			"label": __("Place Type"),
			"fieldtype": "MultiSelectList",
			"options": "Place Type",
			get_data: function(txt) {
				return frappe.db.get_link_options('Place Type', txt);
			}
		}
	]
};