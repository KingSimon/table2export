/*
 * 作者: kingsimon 2015-12-30
 * 说明:依赖jquery
 *
 * 方法名:table2export
 * 方法说明:导出表格(支持格式xls、doc、csv、json、txt)
 * 调用:
 * $("#datatable").table2export({
 *   exclude: ".no-export",                               //排除导出项
 *   style: ["color", "backgroundColor", "fontWeight"],   //导出xls或doc支持样式
 *   name: "Excel Document Name",                         //
 *   filename: "myFileName",                              //导出文件名
 *   type:"doc"                                           //导出文件格式
 * });
 * 参数说明:
 * @param 自定义参数对象
 *
 */
;
(function($) {
  var pluginName = "table2export",

      defaults = {
        exclude: ".no-export",
        style: ["color", "backgroundColor", "fontWeight", "width", "border", 'border-left', 'border-top', 'border-right', 'border-bottom'],
        name: "Table2Export",
        filename: "fileName",
        type: "xls"
      };

  // The actual plugin constructor
  function Plugin(element, options) {
    this.element = element;
    // jQuery has an extend method which merges the contents of two or
    // more objects, storing the result in the first object. The first object
    // is generally empty as we don't want to alter the default options for
    // future instances of the plugin
    //
    this.settings = $.extend({}, defaults, options);
    this.defaults = defaults;
    this.name = pluginName;
    this.typeMap = {
      json: this.toJson,
      txt: this.toTxt,
      csv: this.toCSV,
      doc: this.toDoc,
      xls: this.toExcel
    };
    this.uri = {
      json: 'data:application/json;base64,',
      txt: 'data:csv/txt;charset=utf-8,',
      csv: 'data:csv/txt;charset=utf-8,\ufeff',
      doc: 'data:application/vnd.ms-doc;base64,',
      xls: 'data:application/vnd.ms-excel;base64,',
    };

    this.base64 = function(s) {
      return window.btoa(unescape(encodeURIComponent(s)));
    };
    this.format = function(s, c) {
      return s.replace(/{(\w+)}/g, function(m, p) {
        return c[p];
      });
    };
    this.init();
  }

  Plugin.prototype = {
    init: function() {
      var e = this;
      var type = e.settings.type;
      if (e.typeMap[type]) {
        e.typeMap[type].call(e);
      } else {
        console.log('不支持当前格式导出文件');
      }
    },
    toJson: function() {

      var e = this,
          fullTemplate = "",
          link, a;

      var jsonExportArray = [];
      $(e.element).each(function(i, o) {
        var $table = recoverTable(o, e.settings);
        var jsonHeaderArray = [];
        $table.find('thead').find('tr').each(function() {
          $(this).find('th').each(function(index) {
            var resultText = $.trim($(this).text());
            var $input = $(this).find('input');
            var $select = $(this).find('select');
            if($input.length>0){
              resultText = $.trim($input.val());
            }else if($select.length>0){
              resultText = $.trim($select.find('option:selected').text());
            }
            jsonHeaderArray.push(resultText);
          });
        });
        var jsonArray = [];
        $table.find('tbody').find('tr').each(function(i) {
          $(this).find('td').each(function() {
            if (!jsonArray[i]) jsonArray[i] = [];
            var resultText = $.trim($(this).text());
            var $input = $(this).find('input');
            var $select = $(this).find('select');
            if($input.length>0){
              resultText = $.trim($input.val());
            }else if($select.length>0){
              resultText = $.trim($select.find('option:selected').text());
            }
            jsonArray[i].push(resultText);
          });
        });
        jsonExportArray.push({
          header: jsonHeaderArray,
          data: jsonArray
        });
      });

      fullTemplate = JSON.stringify(jsonExportArray);
      e.toExport(fullTemplate);
    },
    toTxt: function() {
      var e = this;
      e.toCSV();
    },
    toCSV: function() {
      var e = this,
          fullTemplate = "";
      var data = "";
      $(e.element).each(function(index, obj) {
        var $table = recoverTable(obj, e.settings);
        var table = $table[0];

        for (var i = 0, row; row = table.rows[i]; i++) {
          for (var j = 0, col; col = row.cells[j]; j++) {
            data = data + (j ? ',' : '') + fixCSVField(col.innerHTML);
          }
          data = data + "\r\n";
        }
        data = data + "\r\n";
      });

      fullTemplate = data;
      e.toExport(fullTemplate);
    },
    toDoc: function() {
      var e = this;
      e.toExcel();
    },
    toExcel: function() {
      var e = this;

      e.template = {
        head: "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns=\"http://www.w3.org/TR/REC-html40\"><head><meta charset=\"utf-8\" /><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>",
        sheet: {
          head: "<x:ExcelWorksheet><x:Name>",
          tail: "</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>"
        },
        mid: "</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body>",
        table: {
          head: "<table style=\"border-collapse: collapse;border-spacing: 0;border: 2px solid #ddd;\">",
          tail: "</table>"
        },
        foot: "</body></html>"
      };

      e.tableRows = [];

      // get contents of table except for exclude
      $(e.element).each(function(i, o) {
        var tempRows = "";
        $(o).find("tr").not(e.settings.exclude).each(function(i, o) {
          var tempTd = "";
          $(o).find("td,th").not(e.settings.exclude).each(function(k, j) {
            var tag = $(j).get(0).tagName;
            var colspan = $(j).attr("colspan") ? " colspan=\"" + $(j).attr("colspan") + "\"" : "";
            var rowspan = $(j).attr("rowspan") ? " rowspan=\"" + $(j).attr("rowspan") + "\"" : "";
            var styleArry = e.settings.style;
            var style = "";
            $.each(styleArry, function(index, item) {
              style += item + ":" + $(j).css(item) + ";";
            });
            var resultText = $.trim($(this).text());
            var $input = $(j).find('input');
            var $select = $(j).find('select');
            if($input.length>0){
              resultText = $.trim($input.val());
            }else if($select.length>0){
              resultText = $.trim($select.find('option:selected').text());
            }
            tempTd += "<" + tag + colspan + rowspan + " style=\"" + style + "\">" + resultText + "</" + tag + ">";
          });
          tempRows += "<tr>" + tempTd + "</tr>";
        });
        e.tableRows.push(tempRows);
      });


      e.tableToExcel(e.tableRows, e.settings.name);
    },
    tableToExcel: function(table, name) {
      var e = this,
          fullTemplate = "";

      e.ctx = {
        worksheet: name || "Worksheet",
        table: table
      };

      fullTemplate = e.template.head;

      if ($.isArray(table)) {
        $.each(table, function(i, o) {
          //fullTemplate += e.template.sheet.head + "{worksheet" + i + "}" + e.template.sheet.tail;
          fullTemplate += e.template.sheet.head + "Table" + i + "" + e.template.sheet.tail;
        });
      }

      fullTemplate += e.template.mid;

      if ($.isArray(table)) {
        $.each(table, function(i, o) {
          fullTemplate += e.template.table.head + "{Table" + i + "}" + e.template.table.tail;
        });
      }

      fullTemplate += e.template.foot;

      $.each(table, function(i, o) {
        e.ctx["Table" + i] = table[i];
      });
      delete e.ctx.table;
      fullTemplate = e.format(fullTemplate, e.ctx);
      e.toExport(fullTemplate);
      return true;
    },
    toExport: function(fullTemplate) {
      var e = this;

      var msie = window.navigator.userAgent.indexOf("MSIE");
      if (msie > 0 || !!window.navigator.userAgent.match(/Trident.*rv\:11\./)) // If Internet Explorer
      {
        if (typeof Blob !== "undefined") {
          if (e.settings.type == 'csv') {
            fullTemplate = ['\ufeff' + fullTemplate]
          } else {
            fullTemplate = [fullTemplate];
          }
          var blob1 = new Blob(fullTemplate, {
            type: "text/html,charset=utf-8"
          });
          window.navigator.msSaveBlob(blob1, getFileName(e.settings));
        } else {
          //otherwise use the iframe and save
          //requires a blank iframe on page called txtArea1
          if ($('#txtArea1').length === 0)
            $("<iframe width='600' height='600' id='txtArea1' style='display:none'></iframe>").prependTo('body');
          txtArea1.document.open("text/html,charset=utf-8", "replace");
          if (e.settings.type == 'csv') {
            txtArea1.document.write('\ufeff' + fullTemplate);
          } else {
            txtArea1.document.write(fullTemplate);
          }
          txtArea1.document.close();
          txtArea1.focus();
          sa = txtArea1.document.execCommand("SaveAs", true, getFileName(e.settings));
          $('#txtArea1').remove();
        }

      } else {
        var link, a;
        if (e.settings.type == 'csv' || e.settings.type == 'txt') {
          if (e.settings.type == 'csv')
            fullTemplate = '\ufeff' + fullTemplate;
          var blob = new Blob([fullTemplate], {
            type: 'text/csv,charset=utf-8'
          });
          link = URL.createObjectURL(blob);
        } else {
          link = e.uri[e.settings.type] + e.base64(fullTemplate);
        }

        a = document.createElement("a");
        a.download = getFileName(e.settings);
        a.href = link;
        a.click();
      }
    }
  };

  function recoverTable(table, settings) {
    var $table = $(table).clone();
    //清除隐藏项和排除列
    $.each($table.find('th,tr,td'), function() {
      if ($(this).css('display') == "none") {
        $(this).remove();
      }
    });
    $table.find(settings.exclude).remove();
    //恢复列数
    $.each($table.find('tbody>tr'), function(x) {
      $.each($(this).find('td'), function(y) {
        var colspan = Number($(this).attr('colspan')) || 1;
        if (colspan > 1) {
          for (var k = 0; k < colspan - 1; k++) {
            var $obj = $(this).clone(true).removeAttr('colspan').removeAttr('rowspan');
            $(this).after($obj);
            $obj.addClass('hide');
          }
        }
      });
    });
    //恢复行数
    var list = $.map($table.find('tbody>tr'), function(elem, index) {
      return $(elem).find('td').length;
    });
    var len1 = Math.max.apply(null, list);
    var len2 = $table.find('tbody>tr').length;
    for (var i = 0; i < len1; i++) {
      for (var j = 0; j < len2; j++) {
        var $item = $table.find('tbody>tr:nth-child(' + (j + 1) + ')>td:nth-child(' + (i + 1) + ')');
        var rowspan = Number($item.attr('rowspan')) || 1;
        if (rowspan > 1) {
          for (var k = 0; k < rowspan - 1; k++) {
            var $obj = $item.clone(true).removeAttr('colspan').removeAttr('rowspan');
            var $after = $table.find('tbody>tr:nth-child(' + (j + k + 2) + ')>td:nth-child(' + (i + 1) + ')');
            $after.before($obj);
            $obj.addClass('hide');
          }
        }
      }
    }

    //还原项
    $.each($table.find('tbody>tr>td'), function() {
      $(this).removeAttr('colspan').removeAttr('rowspan').removeClass('hide');
    });

    return $table;
  }

  function getFileName(settings) {
    return (settings.filename ? settings.filename : "table2export") + "." + settings.type;
  }

  function fixCSVField(value) {
    var fixedValue = value;
    var addQuotes = (value.indexOf(',') !== -1) || (value.indexOf('\r') !== -1) || (value.indexOf('\n') !== -1);
    var replaceDoubleQuotes = (value.indexOf('"') !== -1);

    if (replaceDoubleQuotes) {
      fixedValue = fixedValue.replace(/"/g, '""');
    }
    if (addQuotes || replaceDoubleQuotes) {
      fixedValue = '"' + fixedValue + '"';
    }
    return fixedValue;
  }

  $.fn[pluginName] = function(options) {
    var e = this;
    var type = options.type || "xls";
    var plugin = new Plugin(this, options);
    e.each(function(index, elem) {
      $.data(elem, "plugin_" + pluginName + "_" + type, plugin);
    });
    // chain jQuery functions
    return e;
  };


})(jQuery)
