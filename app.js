var createError = require('http-errors');
var express = require('express');
var path = require('path');
var cookieParser = require('cookie-parser');
var logger = require('morgan');

var indexRouter = require('./routes/index');
var usersRouter = require('./routes/users');


const XlsxPopulate = require('xlsx-populate');

var app = express();

// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'hbs');

app.use(logger('dev'));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

app.use('/', indexRouter);
app.use('/users', usersRouter);


app.get('/excel', function(req,res,next){
  var t1 = Date.now(),t2,t3,t4;
  XlsxPopulate.fromFileAsync(path.join(__dirname, '/excel/3 PH FRT 7775.xlsm'))
    .then(workbook => {
        t2 = Date.now();
        for(var i=0;i<100;i++){
          workbook.definedName("setno").value(req.query.setno);
        }
        t3 = Date.now();
        return workbook.outputAsync();
    })
    .then(blob =>{
        t4 = Date.now();
        res.writeHead(200, {
          'Content-disposition': 'attachment;filename=3 PH FRT 7775  '+req.query.setno+'.xlsm'
        });
        console.log(t2-t1,t3-t2,t4-t3,t4-t1);
        res.end(blob);
    });
});


// catch 404 and forward to error handler
app.use(function(req, res, next) {
  next(createError(404));
});

// error handler
app.use(function(err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get('env') === 'development' ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render('error');
});

module.exports = app;
