function _purl(u) {
	var f = xrep(xrep(u, "%%REP%%", rPath2Root), "\\", "/").split("/");
	var i = 0;
	while(i<f.length) {
		if(f[i] == "..") {
			for(var j=i-1; j<f.length-2; j++)
				f[j] = f[j+2];
			f.pop();
			f.pop();
		}
		i++;
	}
	return f.join("/");
}