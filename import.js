/**
 * Created by Eric on 17/11/15.
 */

function readFile(obj) {
    var file = obj.files[0];
    var reader = new FileReader();

    reader.readAsBinaryString(file);
    reader.onload = function (e) {
        var data = e.target.result;
        var wb = XLSX.read(data, {type: "binary"});

        //此处对excel数据进行处理
        console.log(wb);
    };
}