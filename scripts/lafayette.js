jQuery(document).ready(function() {
    //Responsive Tabs
    if (jQuery(".resp-tabs-list").length > 0) {
        jQuery("#service-information").easyResponsiveTabs({
            type: 'default', //Types: default, vertical, accordion           
            width: 'auto', //auto or any custom width
            fit: true,   // 100% fits in a container
            closed: false, // Close the panels on start, the options 'accordion' and 'tabs' keep them closed in there respective view types
            activate: function () { }  // Callback function, gets called if tab is switched
        });
    }
});

function changeClass(id, newClass) {
		 var elemToChange = document.getElementById(id);
		 elemToChange.className = newClass;
      }

function isEmail(string) {
	if (!string) return false;

		var iChars = "*|,\":<>[]{}`\';()&$#%";
		for (var i = 0; i < string.length; i++){
			if (iChars.indexOf(string.charAt(i)) != -1){
				return false;
			}
		}
	return true;
}                      
				   
function emailCheck(myForm){	
	//alert("form alert " + myForm);
				
	if (isEmail(myForm.email.value) == false){
		alert("Please enter a valid email address.");
		myForm.email.focus();   
	return false;
	}
}

function load() {
	if (GBrowserIsCompatible()) {
       var map = new GMap2(document.getElementById("map"));
	map.addControl(new GSmallMapControl());
	map.addControl(new GMapTypeControl());
	map.setCenter(new GLatLng(43.084287, -87.907485), 13);
	map.openInfoWindow(map.getCenter(),
                  document.createTextNode("3710 North Richards Street, Milwaukee, WI 53212"));
	
	var baseIcon = new GIcon();
	baseIcon.shadow = "http://labs.google.com/ridefinder/images/mm_20_shadow.png";
	baseIcon.iconSize = new GSize(20, 34);
	baseIcon.shadowSize = new GSize(22, 20);
	baseIcon.iconAnchor = new GPoint(6, 20);
	baseIcon.infoWindowAnchor = new GPoint(5, 1);

	// Creates a numbered marker
function createMarker(point, index) {
	var number = String.fromCharCode("1".charCodeAt(0) + index);
	var icon = new GIcon(baseIcon);
		icon.image = "http://www.lafayettetesting.com/images/markers/blank.png";
		var marker = new GMarker(point, icon);

		return marker;
	}
	for (var i = 0; i < 1; i++) {
 			map.addOverlay(createMarker(new GLatLng(43.084287, -87.907485), i));
	}
	
	// Download the data in data.xml and load it on the map. The format we
	// expect is:
	// <markers>
	//   <marker lat="37.441" lng="-122.141"/>
	//   <marker lat="37.322" lng="-121.213"/>
	// </markers>
	GDownloadUrl("data.xml", function(data, responseCode) {
	  var xml = GXml.parse(data);
	  var markers = xml.documentElement.getElementsByTagName("marker");
	  for (var i = 0; i < markers.length; i++) {
	    var point = new GLatLng(parseFloat(markers[i].getAttribute("lat")),
	                            parseFloat(markers[i].getAttribute("lng")));
	    map.addOverlay(new GMarker(point));
	  }
	});
     }
}
	
function initAll(){
	load();
}