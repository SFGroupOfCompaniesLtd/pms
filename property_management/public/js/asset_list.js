frappe.listview_settings['Asset'] = {
	refresh: function(frm) {
        $('*[data-filter="property_status,=,Closed"]').css({'color':'#e24c4c', 'background':'#fff5f5'}); // red
        $('*[data-filter="property_status,=,Rent/Lease"]').css({'color':'#f8814f', 'background':'#fff5f0'}); // orange
        $('*[data-filter="property_status,=,Booked"]').css({'color':'#1579d0', 'background':'#D3E9FA'}); // blue
        $('*[data-filter="property_status,=,Available"]').css({'color':'#2f9d58', 'background':'#eaf5ee'}); // green
	}
};