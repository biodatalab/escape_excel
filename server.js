var express     = require("express");
var app         = express();
var multer      = require('multer');
var http        = require("http");
var upload      = multer();
var cons        = require("consolidate");
var spawn       = require("child_process").spawn;
var session     = require('express-session');

app.engine('html', cons.mustache);
app.set('view engine', 'html');
app.set('views', __dirname + '/server/views');

app.use("/static", express.static(__dirname + "/server/static"));

app.get("/", function(req, res) {
    res.render("index");
});

// used to add command line arguments to
// the escape excel program
function argsBuilder(args, body) {
    return {
        enable: function(flag) {
            if (body[flag]) {
                args.push("--" + flag);
            }
        }
    }
}

app.post("/upload", upload.single('file'), function(req, res) {
    if (req.file) {
        var args = ["escape_excel.pl"];
        var builder = argsBuilder(args, req.body);
        builder.enable("no-dates");
        builder.enable("no-sci");
        builder.enable("no-zeros");
        builder.enable("paranoid");

        var escape = spawn("perl", args, {
            detached: true,
            stdio: 'pipe'
        });
        res.setHeader('Content-disposition', 'attachment; filename=' + req.file.originalname);
        res.setHeader('Content-type', "text/plain");
        escape.stdout.pipe(res);
        escape.stdin.end(req.file.buffer);
    } else {
        res.redirect("/");
    }
});

http.createServer(app).listen(8000);
console.log("listening at :8000");