# NAME

node-win32com - Asynchronous, non-blocking win32com ( win32ole / win32api ) wrapper and tools for [node.js](https://github.com/joyent/node) .

win32com ( win32ole / win32api ) makes accessibility from node.js to Excel, Word, Access, Outlook, InternetExplorer and so on. It does'nt need TypeLibrary.


# USAGE

Install with `npm install win32com`.

It works as... (version 0.1.x)

``` js
var win32com = require('win32com');
var xl = win32com.client.Dispatch('Excel.Application', 'C'); // locale
xl.Visible = true;
var book = xl.Workbooks.Add();
var sheet = book.Worksheets(1);
sheet.Name = 'sheetnameA utf8';
sheet.Cells(1, 2).Value = 'test utf8';
var rg = sheet.Range(sheet.Cells(2, 2), sheet.Cells(4, 4));
rg.RowHeight = 5.18;
rg.ColumnWidth = 0.58;
rg.Interior.ColorIndex = 6; // Yellow
book.SaveAs('testfileutf8.xls');
xl.ScreenUpdating = true;
xl.Workbooks.Close();
xl.Quit();
```

But now it implements as... (version 0.0.x)

``` js
win32com.client = new win32com.Client;
try{
  var win32com = require('win32com');
  var xl = win32com.client.Dispatch('Excel.Application', 'C'); // locale
  xl.set('Visible', true);
  var book = xl.get('Workbooks').call('Add');
  var sheet = book.get('Worksheets', [1]);
  try{
    sheet.set('Name', 'sheetnameA utf8');
    sheet.get('Cells', [1, 2]).set('Value', 'test utf8');
    var rg = sheet.get('Range',
      [sheet.get('Cells', [2, 2]), sheet.get('Cells', [4, 4])]);
    rg.set('RowHeight', 5.18);
    rg.set('ColumnWidth', 0.58);
    rg.get('Interior').set('ColorIndex', 6); // Yellow
    var result = book.call('SaveAs', ['testfileutf8.xls']);
    console.log(result);
  }catch(e){
    console.log('(exception cached)\n' + e);
  }
  xl.set('ScreenUpdating', true);
  xl.get('Workbooks').call('Close');
  xl.call('Quit');
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
win32com.client.Finalize(); // must be called (version 0.0.x)
```


# FEATURES

* So much implements.
* Implement accessors getter, setter and caller.
* npm


# API

See the [API documentation](https://github.com/idobatter/node-win32com/wiki) in the wiki.


# BUILDING

This project uses VC++ 2008 Express (or later) and Python 2.6 (or later) .
(When using Python 2.5, it needs [multiprocessing 2.5 back port](http://pypi.python.org/pypi/multiprocessing/) .)

Bulding also requires node-gyp to be installed. You can do this with npm:

    npm install -g node-gyp

To obtain and build the bindings:

    git clone git://github.com/idobatter/node-win32com.git
    cd node-win32com
    node-gyp configure
    node-gyp build

You can also use [`npm`](https://github.com/isaacs/npm) to download and install them:

    npm install win32com


# TESTS

[mocha](https://github.com/visionmedia/mocha) is required to run unit tests.

    npm install -g mocha
    nmake /a test


# CONTRIBUTORS

* [idobatter](https://github.com/idobatter)


# ACKNOWLEDGEMENTS

Inspired [pywin32](http://pypi.python.org/pypi/pywin32)


# LICENSE

`node-win32com` is [BSD licensed](https://github.com/idobatter/node-win32com/raw/master/LICENSE).
