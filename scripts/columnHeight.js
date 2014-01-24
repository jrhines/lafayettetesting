//Dynamically set column heights to match. div ID's must be set to "leftColumn" and "rightColumn". 
//must be nested in a div with the ID columnContainer
//10px height to bottom for spacing. 

function columnHeight() {
	if (document.getElementById('leftColumn') != null && document.getElementById('rightColumn') != null){
		document.getElementById('leftColumn').style.height=
		document.getElementById('rightColumn').style.height=
		(document.getElementById('columnContainer').offsetHeight)+"px";
	}
}

window.onload=columnHeight;