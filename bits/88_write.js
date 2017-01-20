function write_zip_type(wb, opts) {
	var o = opts||{};
	var z = write_zip(wb, o);
	var type;
	switch(o.type) {
		case "base64":
			type="base64";
		break;
		case "binary":
			type="binary";
		break;
		case "buffer":
		case "file":
			type="nodebuffer";
		break;
		default: throw new Error("Unrecognized type " + o.type);
	}
	var params = Object.assign({},o,{type:type});
	if(o.type != "file")
		return z.generate(params);
	return _fs.writeFileSync(o.file, z.generate(params));
}

function writeSync(wb, opts) {
	var o = opts||{};
	switch(o.bookType) {
		case 'xml': return write_xlml(wb, o);
		default: return write_zip_type(wb, o);
	}
}

function writeFileSync(wb, filename, opts) {
	var o = opts||{}; o.type = 'file';
	o.file = filename;
	switch(o.file.substr(-5).toLowerCase()) {
		case '.xlsx': o.bookType = 'xlsx'; break;
		case '.xlsm': o.bookType = 'xlsm'; break;
		case '.xlsb': o.bookType = 'xlsb'; break;
	default: switch(o.file.substr(-4).toLowerCase()) {
		case '.xls': o.bookType = 'xls'; break;
		case '.xml': o.bookType = 'xml'; break;
	}}
	return writeSync(wb, o);
}
