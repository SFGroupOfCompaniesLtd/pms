frappe.ui.form.on("Asset", "address", function(frm, cdt, cdn) {
    if (frm.doc.address){
        return frm.call({
            method: "frappe.contacts.doctype.address.address.get_address_display",
            args: {
                "address_dict": frm.doc.address
            },
            callback: function(r) {
                if(r.message)
                frm.set_value("address_display", r.message);
            }
        });
    }
    else {
        frm.set_value("address_display", "");
    }
});

frappe.ui.form.on("Asset", {
    setup: function(frm) {
        frm.set_query("address", function() {
            return {
                filters: [
                    ["Dynamic Link","link_doctype", "=", "Customer"],
                    ["Dynamic Link","link_name", "=", frm.doc.property_manager]
                ]
            };
        });
    }
});


frappe.ui.form.on("Asset", {
    setup: function(frm) {
        frm.set_query("contact", function() {
            return {
                filters: [
                    ["Dynamic Link","link_doctype", "=", "Customer"],
                    ["Dynamic Link","link_name", "=", frm.doc.property_manager]
                ]
            };
        });
    }
});

frappe.ui.form.on("Asset", {
    before_save: function(frm) {
        frm.set_value('gfa_m', frm.doc.gfa_sqft * 0.092903);
        frm.set_value('total_price', frm.doc.gfa_sqft * frm.doc.unit_price);
    }
});


frappe.ui.form.on('Property Insurance', {
    form_render(frm, cdt, cdn){
        frm.fields_dict.property_insurance.grid.wrapper.find('.btn-attach').addClass("btn-primary").removeClass("btn-default");
    }
});


frappe.ui.form.on('Floor Plan', {
    form_render(frm, cdt, cdn){
        frm.fields_dict.floor_plan.grid.wrapper.find('.btn-attach').addClass("btn-primary").removeClass("btn-default");
    }
});


frappe.ui.form.on('Photos', {
    form_render(frm, cdt, cdn){
        frm.fields_dict.photos.grid.wrapper.find('.btn-attach').addClass("btn-primary").removeClass("btn-default");
    }
});


frappe.ui.form.on('Documents', {
    form_render(frm, cdt, cdn){
        frm.fields_dict.documents.grid.wrapper.find('.btn-attach').addClass("btn-primary").removeClass("btn-default");
    }
});


frappe.ui.form.on("Asset", "before_save", function(frm, cdt, cdn) {
    $.each(frm.doc.finance_books || [], function(i, d) {
        var end_date = frappe.datetime.add_months(d.depreciation_start_date, (d.total_number_of_depreciations * d.frequency_of_depreciation) - 1);
        frm.set_value('start_date', d.depreciation_start_date);
        frm.set_value('end_date', end_date);
    });
});


frappe.ui.form.on('Asset', {
    rent_type: function(frm) {
        frm.set_value('calculate_depreciation', 1);
        if (frm.doc.opening_accumulated_depreciation === 0) {
            frm.clear_table('finance_books');
            var d = frm.add_child("finance_books");
            d.depreciation_method="Manual";
            if (frm.doc.rent_type == "Monthly") {
                d.total_number_of_depreciations = 12;
                d.frequency_of_depreciation = 1;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 1);
                d.expected_value_after_useful_life = 0;
            }
            else if (frm.doc.rent_type == "Quarterly") {
                d.total_number_of_depreciations = 4;
                d.frequency_of_depreciation = 3;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 3);
                d.expected_value_after_useful_life = 0;
            }
            else if (frm.doc.rent_type == "Half-yearly") {
                d.total_number_of_depreciations = 2;
                d.frequency_of_depreciation = 6;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 6);
                d.expected_value_after_useful_life = 0;
            }
            else if (frm.doc.rent_type == "Yearly") {
                d.total_number_of_depreciations = 1;
                d.frequency_of_depreciation = 12;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 12);
                d.expected_value_after_useful_life = 0;
            }
            frm.refresh_fields("finance_books");
        }
        else if (frm.doc.opening_accumulated_depreciation !== 0) {
            frm.clear_table('finance_books');
            var d = frm.add_child("finance_books");
            d.depreciation_method="Manual";
            if (frm.doc.rent_type == "Monthly") {
                d.total_number_of_depreciations = 12;
                d.frequency_of_depreciation = 1;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 1);
                d.expected_value_after_useful_life = frm.doc.opening_accumulated_depreciation;
            }
            else if (frm.doc.rent_type == "Quarterly") {
                d.total_number_of_depreciations = 4;
                d.frequency_of_depreciation = 3;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 3);
                d.expected_value_after_useful_life = frm.doc.opening_accumulated_depreciation;
            }
            else if (frm.doc.rent_type == "Half-yearly") {
                d.total_number_of_depreciations = 2;
                d.frequency_of_depreciation = 6;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 6);
                d.expected_value_after_useful_life = frm.doc.opening_accumulated_depreciation;
            }
            else if (frm.doc.rent_type == "Yearly") {
                d.total_number_of_depreciations = 1;
                d.frequency_of_depreciation = 12;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 12);
                d.expected_value_after_useful_life = frm.doc.opening_accumulated_depreciation;
            }
            frm.refresh_fields("finance_books");
        }
    },

    available_for_use_date: function(frm) {
        if (frm.doc.calculate_depreciation === 1 && frm.doc.opening_accumulated_depreciation === 0) {
            frm.clear_table('finance_books');
            var d = frm.add_child("finance_books");
            d.depreciation_method="Manual";
            if (frm.doc.rent_type == "Monthly") {
                d.total_number_of_depreciations = 12;
                d.frequency_of_depreciation = 1;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 1);
                d.expected_value_after_useful_life = 0;
            }
            else if (frm.doc.rent_type == "Quarterly") {
                d.total_number_of_depreciations = 4;
                d.frequency_of_depreciation = 3;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 3);
                d.expected_value_after_useful_life = 0;
            }
            else if (frm.doc.rent_type == "Half-yearly") {
                d.total_number_of_depreciations = 2;
                d.frequency_of_depreciation = 6;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 6);
                d.expected_value_after_useful_life = 0;
            }
            else if (frm.doc.rent_type == "Yearly") {
                d.total_number_of_depreciations = 1;
                d.frequency_of_depreciation = 12;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 12);
                d.expected_value_after_useful_life = 0;
            }
            frm.refresh_fields("finance_books");
        }
        else if (frm.doc.calculate_depreciation === 1 && frm.doc.opening_accumulated_depreciation !== 0) {
            frm.clear_table('finance_books');
            var d = frm.add_child("finance_books");
            d.depreciation_method="Manual";
            if (frm.doc.rent_type == "Monthly") {
                d.total_number_of_depreciations = 12;
                d.frequency_of_depreciation = 1;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 1);
                d.expected_value_after_useful_life = frm.doc.opening_accumulated_depreciation;
            }
            else if (frm.doc.rent_type == "Quarterly") {
                d.total_number_of_depreciations = 4;
                d.frequency_of_depreciation = 3;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 3);
                d.expected_value_after_useful_life = frm.doc.opening_accumulated_depreciation;
            }
            else if (frm.doc.rent_type == "Half-yearly") {
                d.total_number_of_depreciations = 2;
                d.frequency_of_depreciation = 6;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 6);
                d.expected_value_after_useful_life = frm.doc.opening_accumulated_depreciation;
            }
            else if (frm.doc.rent_type == "Yearly") {
                d.total_number_of_depreciations = 1;
                d.frequency_of_depreciation = 12;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 12);
                d.expected_value_after_useful_life = frm.doc.opening_accumulated_depreciation;
            }
            frm.refresh_fields("finance_books");
        }
    },

    opening_accumulated_depreciation: function(frm) {
        if (frm.doc.calculate_depreciation === 1 && frm.doc.opening_accumulated_depreciation === 0) {
            frm.clear_table('finance_books');
            var d = frm.add_child("finance_books");
            d.depreciation_method="Manual";
            if (frm.doc.rent_type == "Monthly") {
                d.total_number_of_depreciations = 12;
                d.frequency_of_depreciation = 1;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 1);
                d.expected_value_after_useful_life = 0;
            }
            else if (frm.doc.rent_type == "Quarterly") {
                d.total_number_of_depreciations = 4;
                d.frequency_of_depreciation = 3;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 3);
                d.expected_value_after_useful_life = 0;
            }
            else if (frm.doc.rent_type == "Half-yearly") {
                d.total_number_of_depreciations = 2;
                d.frequency_of_depreciation = 6;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 6);
                d.expected_value_after_useful_life = 0;
            }
            else if (frm.doc.rent_type == "Yearly") {
                d.total_number_of_depreciations = 1;
                d.frequency_of_depreciation = 12;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 12);
                d.expected_value_after_useful_life = 0;
            }
            frm.refresh_fields("finance_books");
        }
        else if (frm.doc.calculate_depreciation === 1 && frm.doc.opening_accumulated_depreciation !== 0) {
            frm.clear_table('finance_books');
            var d = frm.add_child("finance_books");
            d.depreciation_method="Manual";
            if (frm.doc.rent_type == "Monthly") {
                d.total_number_of_depreciations = 12;
                d.frequency_of_depreciation = 1;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 1);
                d.expected_value_after_useful_life = frm.doc.opening_accumulated_depreciation;
            }
            else if (frm.doc.rent_type == "Quarterly") {
                d.total_number_of_depreciations = 4;
                d.frequency_of_depreciation = 3;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 3);
                d.expected_value_after_useful_life = frm.doc.opening_accumulated_depreciation;
            }
            else if (frm.doc.rent_type == "Half-yearly") {
                d.total_number_of_depreciations = 2;
                d.frequency_of_depreciation = 6;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 6);
                d.expected_value_after_useful_life = frm.doc.opening_accumulated_depreciation;
            }
            else if (frm.doc.rent_type == "Yearly") {
                d.total_number_of_depreciations = 1;
                d.frequency_of_depreciation = 12;
                d.depreciation_start_date = frappe.datetime.add_months(frm.doc.available_for_use_date, 12);
                d.expected_value_after_useful_life = frm.doc.opening_accumulated_depreciation;
            }
            frm.refresh_fields("finance_books");
        }
    }    
    
});

frappe.ui.form.on('Asset', {
    refresh(frm) {
        console.log("Test ::::::::::::::::::::::")
        frm.set_df_property('depreciation_schedule_sb', 'hide', 1)
        frm.add_custom_button(__("Available"), function() {
            frm.set_value('property_status', "Available");
            frm.save();
            if (frm.doc.docstatus == 1) {
                frm.save('Update');
            }
        },__("Property Status"));
        frm.add_custom_button(__("Rent/Lease"), function() {
            frm.set_value('property_status', "Rent/Lease");
            frm.save();
            if (frm.doc.docstatus == 1) {
                frm.save('Update');
            }
        },__("Property Status"));
        frm.add_custom_button(__("Booked"), function() {
            frm.set_value('property_status', "Booked");
            frm.save();
            if (frm.doc.docstatus == 1) {
                frm.save('Update');
            }
        },__("Property Status"));
        frm.add_custom_button(__("Closed"), function() {
            frm.set_value('property_status', "Closed");
            frm.save();
            if (frm.doc.docstatus == 1) {
                frm.save('Update');
            }
        },__("Property Status"));

        if (frm.doc.docstatus == 1) {
            frm.page.set_inner_btn_group_as_primary(__("Create"));
            frm.add_custom_button(__("Tenancy"), function() {
                    frappe.route_options = {
                        "asset": frm.doc.name
                    };
                    frappe.set_route("tenancy", "new-tenancy");
            }, __("Create"));
        }
    }
});