/**
 * Created by Eric on 17/11/14.
 */

//导出
document.getElementById('J_export').addEventListener('click', downloadExl);

function downloadExl() {
    //这里的数据是用来定义导出的格式类型
    var wopts = {bookType: 'xlsx', bookSST: true, type: 'binary', cellStyles: true};
    var wb = {SheetNames: ['Sheet1'], Sheets: {}, Props: {}};

    var data = {};

    //数据整合
    data["A1"] = {
        "t": "s",
        "v": "测试1",
        "r": "<t>测试1</t>",
        "h": "测试1",
        "w": "测试1",
        "s": {
            fill: {fgColor: {rgb: "FFFF00"}},
            alignment: {vertical: "center", horizontal: "center", wrapText: "true"},
            font: {sz: 14, bold: true, color: {rgb: "000"}}
        }
    };

    data["A2"] = {
        "t": "s",
        "v": "测试2",
        "r": "<t>测试2</t>",
        "h": "测试2",
        "w": "测试2"
    };

    //表格范围
    data["!ref"] = "A1:IV49";

    //合并数据
    data["!merges"] = [
        {
            //开始
            "s": {
                //开始列
                "c": 0,
                //开始行
                "r": 0
            },
            //结束
            "e": {
                //结束列
                "c": 1,
                //结束行
                "r": 0
            }
        },
        {
            "s": {
                "c": 0,
                "r": 1
            },
            "e": {
                "c": 0,
                "r": 2
            }
        }
    ];

    wb.Sheets['Sheet1'] = data;

    var obj = new Blob([s2ab(XLSX.write(wb, wopts))], {type: "application/octet-stream"}),
        fileName = "测试" + '.' + (wopts.bookType == "biff2" ? "xls" : wopts.bookType);

    saveAs(obj, fileName);
}

function saveAs(obj, fileName) {//当然可以自定义简单的下载文件实现方式
    var tmpa = document.createElement("a");
    tmpa.download = fileName || "下载";
    tmpa.href = URL.createObjectURL(obj); //绑定a标签
    tmpa.click(); //模拟点击实现下载
    setTimeout(function () { //延时释放
        URL.revokeObjectURL(obj); //用URL.revokeObjectURL()来释放这个object URL
    }, 100);
}

function s2ab(s) {
    if (typeof ArrayBuffer !== 'undefined') {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    } else {
        var buf = new Array(s.length);
        for (var i = 0; i != s.length; ++i) buf[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
}
