/**
 * 前端导出Excel文件
 * @param {Boolean} isMulti 表格是否是多级表头 默认值true
 * @param {String} dom 表格ID或包含表格的ID 
 * @param {String} name 表格导出的名称
 * @param {Boolean} isTotal 表格是否有合计行（如果表格头部含有合计行则不用配置） 默认值false
 * @param {Number} totalNum 表格含有合计字符的个数  默认值 2
 * @param {String} type 导出文件的格式  默认值 xlsx
 * @param {Number} wpx Excel的宽列  默认值 60
 * @param {Boolean} hasRowNum 表格是否需要行号  默认值 false
 * @param {Object} cellWpx 表格自定义列宽  默认值 {0:{wpx: 120}}
 */

function ExportExcel () {

  this.table = ({isMulti = true, dom, name, isTotal = false, totalNum = 2, type = 'xlsx', wpx = 60, hasRowNum = false, cellWpx = {}}) => {
    const _dom = this.setTableDom(isMulti, dom, name, hasRowNum);

    this.download({dom: _dom, name, isTotal, totalNum, type, wpx, cellWpx});
  }

  // 表格DOM组装
  this.setTableDom = (isMulti, dom, name, hasRowNum) => {
    $('.reportTable').remove();
    // 获取表格表头与表格体
    const $title = $(dom).find('.ui-jqgrid-hdiv').eq(0).find('.ui-jqgrid-htable');
    const $body = $(dom).find(".ui-jqgrid-bdiv").eq(0).find('.ui-jqgrid-btable');

    let table_all = $title.clone();
    // 获取表格合并行的数值
    const len = this.getTableThRowspan(table_all.find('tr')[0]);
    //表中
    const table_contentAll = $body.find('tbody').clone();

    if (isMulti) {
      $(table_all.find('tr')[0]).remove();
    }    

    if (hasRowNum) {
      table_all.find('tr').each(function(index, elem) {
        if ($(elem).hasClass('total')) {
          $(elem)
            .find('td')
            .eq(0)
            .remove();
        }

        let hasRN = $(elem).find('th').eq(0).text() === '行号' || $(elem).find('th').eq(0).text() === '序号';
        
        if (hasRN) {            
          $(elem).find('th').eq(0).remove();
        }
      });
    }

    // 添加表格标题
    table_all.find('thead').prepend(`<tr><th colspan=${len}>${name}</th></tr>`);

    // 删除表头隐藏列
    table_all.find('th').each(function(index, elem) {
      if ($(elem).css('display') == 'none') {
        $(elem).remove();
      }
    });
    table_all.find('td').each(function(index, elem) {
      if ($(elem).css('display') == 'none') {
        $(elem).remove();
      }
    });

    // 删除表身隐藏列
    $(table_contentAll.find('tr')[0]).remove();

    table_contentAll.find('td').each(function(index, elem) {
      if ($(elem).css('display') == 'none') {
        $(elem).remove();
      }
    });

    if (hasRowNum) {
      table_contentAll.find('tr').each(function(index, elem) {
        $(elem)
          .find('td')
          .eq(0)
          .remove();
      });
    }

    const table_content = table_contentAll;

    let hasThbody = table_all.find('tbody').length;
    !hasThbody && table_all.append(`<tbody></tbody>`);
    table_all.find('tbody').append(table_content.find('tr'));
  
    table_all.removeClass();
    table_all.attr('class', 'reportTable');
    table_all.css('display', 'none');

    $('body').append(table_all);

    return $('.reportTable')[0];
  }

  // 获取表格列数
  this.getTableThRowspan = tr => {
    $(tr)
      .find('th')
      .each(function(index, elem) {
        if ($(elem).css('display') == 'none') {
          $(elem).remove();
        }
      });
    let len = $(tr).find('th').length;
    return len;
  }

  // 表格数据转换Excel数据
  this.download = ({dom, name, isTotal, totalNum, type, wpx, cellWpx}) => {
    // 表格数据
    const wb = XLSX.utils.table_to_book(dom, { sheet: `${name}`, raw: true });
    console.log(wb);
    const _num = totalNum;
    const _bookType = type;
    const _isTotal = isTotal;
    const _wpx = wpx;
    const wopts = {
      bookType: _bookType,
      bookSST: true,
      type: 'binary',
      cellStyles: true
    };

    const tableData = wb['Sheets'][`${name}`];

    // 设置样式
    this.setExlStyle(tableData, _num, _isTotal, name, _wpx, cellWpx);

    const _wopts = {
      bookType: _bookType,
      bookSST: false,
      type: 'binary'
    };

    const styleWrite = STYLEXLSX.write(wb, _wopts);

    const s2ab = this.s2ab(styleWrite);

    let tmpDown = new Blob(
      [
        s2ab
      ],
      {
        type: ''
      }
    );

    const _name = `${name}.${(wopts.bookType == 'biff2' ? 'xls' : wopts.bookType)}`;
    
    this.saveAs(tmpDown, _name);
  }

  // 设置表格样式
  this.setExlStyle = (data, totalNum, isTotal, name, wpx, cellWpx) => {
    const borderAll = {
      //单元格外侧框线
      top: {
        style: 'thin'
      },
      bottom: {
        style: 'thin'
      },
      left: {
        style: 'thin'
      },
      right: {
        style: 'thin'
      },
      diagonalDown: true
    };
    data['!cols'] = [];
    let totalText = []; // 合计数据
    for (let key in data) {
      if (data[key] instanceof Object) {
        data[key].s = {
          border: borderAll,
          alignment: {
            vertical: 'center', // 垂直居中
            horizontal: 'center', //水平居中对其
            wrapText: true // 自动换行
          }
        };
        data['!cols'].push({ wpx: wpx });
  
        if (data[key].v === '合计') {
          totalText.push(key);
        }
      }
    }

    for (let key in cellWpx) {
      let { wpx } = cellWpx[key];
      data['!cols'][key].wpx = wpx;
    }
  
    if (totalText.length > totalNum) {
      for (let i = totalNum; i < totalText.length; i++) {
        data[totalText[i]].v = '';
      }
    }
  
    if (isTotal) {
      data['!merges'].push({
        s: { r: 3, c: 0 },
        e: { r: 3, c: 1 }
      });
    }

    // 报表头部样式
    if (!data[`A1`]) {
      data[`A1`] = { s: { font: {} }, v: name };
    }

    data[`A1`].s.font = {
      sz: 20,
      bold: true
    };
    data[`A1`].s.alignment = {
      vertical: 'center', // 垂直居中
      horizontal: 'center', //水平居中对其
      wrapText: false // 自动换行
    };

    const deletionTh = this.getDeletionThKey(data);

    this.setDeletionThCell(data, deletionTh, borderAll);

    return data;
  }
  
  // 获取cell
  this.getDeletionThKey = data => {
    const nums = [];
    const letters = [];
    for (let key of Object.keys(data)) {
      if ((data[key].toString() === '[object Object]')) {
        let num = key.replace(/\D+/, '');
        let letter = key.replace(/\d+/, '');
        if (!nums.includes(num)) {
          nums.push(num);
        }
        if (!letters.includes(letter)) {
          letters.push(letter);
        }
      }
    }
    const deletionTh = letters.reduce((arr, v) => {
      for (let num of nums) {
        let o = `${v}${num}`;
        arr.push(o);          
      }
      return arr;
    }, []);
    return deletionTh;
  }

  // 设置空cell
  this.setDeletionThCell = (data, keys, borderAll) => {
    for (let key of keys) {
      if (!data[key]) {
        data[key] = {
          s: {
            border: borderAll
          }
        }
      }
    }
  }

  // 导出文件
  this.saveAs = (obj, fileName) => {
    let tmpa = document.createElement('a');
    tmpa.download = fileName || '下载';
    tmpa.href = URL.createObjectURL(obj);
    // 兼容火狐
    document.body.appendChild(tmpa);
    tmpa.style.display='none';
    tmpa.click();
    setTimeout(function() {
      URL.revokeObjectURL(obj);
    }, 100);
  }

  // 转换数据
  this.s2ab = s => {
    if (typeof ArrayBuffer !== 'undefined') {
      let buf = new ArrayBuffer(s.length);
      let view = new Uint8Array(buf);
      for (let i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
      return buf;
    } else {
      let buf = new Array(s.length);
      for (let i = 0; i != s.length; ++i) buf[i] = s.charCodeAt(i) & 0xff;
      return buf;
    }
  }
}