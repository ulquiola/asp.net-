var ie;
ie = document.all?1:0;

function hL(E){
	if (ie) {
		while (E.tagName!="TR") {
			E=E.parentElement;
		}
	}
	else {
		while (E.tagName!="TR") {
			E=E.parentNode;
		}
	}
	E.className = "H";
}

function dL(E){
	if (ie) {
		while (E.tagName!="TR") {
			E=E.parentElement;
		}
	}
	else {
		while (E.tagName!="TR") {
			E=E.parentNode;
		}
	}
	E.className = "";
}

function CCA(CB){
	if (CB.checked)
		hL(CB);
	else
		dL(CB);
}
