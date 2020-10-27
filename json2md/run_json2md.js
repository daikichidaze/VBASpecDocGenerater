const json2md = require("json2md")

function runJsonToMarkdown(json_string){
    var obj = JSON.parse(json_string)
    return json2md(obj);
}

module.exports = {
	runJsonToMarkdown: runJsonToMarkdown,
};