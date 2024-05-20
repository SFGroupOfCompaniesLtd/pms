// Copyright (c) 2022, Nihantra C. Patel and contributors
// For license information, please see license.txt
/* eslint-disable */

frappe.query_reports["Tenancy Details"] = {
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
			"label": __("Tenancy ID"),
			"fieldtype": "Link",
			"options": "Tenancy",
			get_query: () => {
				return {
					"filters": {
						"docstatus": 1
					}
				}
			}
		},
		{
			"fieldname": "asset",
			"label": __("Property"),
			"fieldtype": "Link",
			"options": "Asset",
			get_query: () => {
				return {
					"filters": {
						"docstatus": 1
					}
				}
			}
		},
		{
			"fieldname": "is_tenant_tenancy",
			"label": __("Is Tenant Tenancy"),
			"fieldtype": "Check"
		},
		{
			"fieldname": "tenant",
			"label": __("Tenant"),
			"fieldtype": "MultiSelectList",
			"options": "Customer",
			"depends_on": "is_tenant_tenancy",
			get_data: function(txt) {
				return frappe.db.get_link_options('Customer', txt, {
					is_tenant: 1
				});
			}
		},
		{
			"fieldname": "is_landlord_tenancy",
			"label": __("Is Landlord Tenancy"),
			"fieldtype": "Check"
		},
		{
			"fieldname": "landlord",
			"label": __("Landlord"),
			"fieldtype": "MultiSelectList",
			"options": "Customer",
			"depends_on": "is_landlord_tenancy",
			get_data: function(txt) {
				return frappe.db.get_link_options('Customer', txt, {
					is_landlord: 1
				});
			}
		},
		{
			"fieldname": "rent_type",
			"label": __("Rent Type"),
			"fieldtype": "Select",
			"options": ["", "Monthly", "Quarterly", "Half-Yearly", "Yearly"]
		},
		{
			"fieldname": "invoice",
			"label": __("Invoice"),
			"fieldtype": "Link",
			"options": "Sales Invoice",
			get_query: () => {
				return {
					"filters": {
						"docstatus": 1
					}
				}
			}
		},
		{
			"fieldname": "is_paid",
			"label": __("Is Paid"),
			"fieldtype": "Check"
		},
		{
			"fieldname": "payment_entry",
			"label": __("Payment Entry"),
			"fieldtype": "Link",
			"options": "Payment Entry",
			"depends_on": "is_paid",
			get_query: () => {
				return {
					"filters": {
						"docstatus": 1
					}
				}
			}
		}
	],
	onload: function(report) {
		report.page.add_inner_button(("Schedule-wise Tenancy Details"), function() {
			frappe.set_route('query-report', 'Schedule-wise Tenancy Details');
		}).addClass("btn-info").removeClass("btn-default");
	}
};