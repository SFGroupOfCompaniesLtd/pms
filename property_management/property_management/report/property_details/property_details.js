// Copyright (c) 2022, Nihantra C. Patel and contributors
// For license information, please see license.txt
/* eslint-disable */

frappe.query_reports["Property Details"] = {
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
			"fieldname": "from_date",
			"label": __("From Date"),
			"fieldtype": "Date",
			"default": frappe.defaults.get_user_default("year_start_date"),
			"reqd": 1
		},
		{
			"fieldname": "to_date",
			"label": __("To Date"),
			"fieldtype": "Date",
			"default": frappe.defaults.get_user_default("year_end_date"),
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
			"fieldname": "location",
			"label": __("Location"),
			"fieldtype": "MultiSelectList",
			"options": "Location",
			get_data: function(txt) {
				return frappe.db.get_link_options('Location', txt);
			}
		},
		{
			"fieldname": "property_manager",
			"label": __("Property Manager"),
			"fieldtype": "MultiSelectList",
			"options": "Customer",
			get_data: function(txt) {
				return frappe.db.get_link_options('Customer', txt, {
					is_property_manager: 1
				});
			}
		},
		{
			"fieldname": "bedrooms",
			"label": __("1/2/3/4/5+ BHK"),
			"fieldtype": "Select",
			"options": ["", "1", "2", "3", "4", "5+"]
		},
		{
			"fieldname": "furnishing",
			"label": __("Furnishing"),
			"fieldtype": "Select",
			"options": ["", "Full Furnished", "Semi Furnished"]
		},
		{
			"fieldname": "property_status",
			"label": __("Property Status"),
			"fieldtype": "Select",
			"options": ["", "Available", "Booked", "Rent/Lease", "Closed"]
		},
		{
			"fieldname": "is_existing_asset",
			"label": __("Is Existing Asset"),
			"fieldtype": "Check"
		}
	]
};