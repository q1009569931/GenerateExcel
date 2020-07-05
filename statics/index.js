var table = document.getElementById("table");
var line = 1;//行数


// 发送数据制作表格
function PostData(id) {
	let rank = document.getElementById("rank").value //精密度等级
	let hole_diff = document.getElementById("hole_diff").value //孔位公差
	let name = id.title; //数据类别
	let num_input = document.getElementById(name).value;
	let data = {
		"name": name,
		"rank": rank,
		"hole_diff": hole_diff,
		"num_input": num_input
	}
	let httpRequest;
	httpRequest = new XMLHttpRequest();
	let csrftoken = getCookie('csrftoken');

	if (!httpRequest) {
		alert('Giving up :( Cannot create an XMLHTTP instance');
		return false;
	}

	//设置处理响应函数
	httpRequest.onreadystatechange = makeResponse;
	httpRequest.open('POST', "/generatorexcel/", true);
	//跨域设置
	httpRequest.setRequestHeader("X-CSRFToken", csrftoken);
	httpRequest.send(JSON.stringify(data));

	function makeResponse() {
		if (httpRequest.readyState == 4 && httpRequest.status == 200){
			var data = JSON.parse(httpRequest.responseText);
			var row = addline(data);
			table.appendChild(row);
		}
	}
}


// 为列表增加新行
function addline(data) {
	var row = document.createElement("tr"); //创建行

	//加入行，填充数据
	var head = document.createElement("td");
	head.innerHTML = line;
	line += 1; //行数加1
	row.appendChild(head);

	var num_require = document.createElement("td");
	num_require.innerHTML = data.num_require;
	row.appendChild(num_require);

	var num_real = document.createElement("td");
	num_real.innerHTML = data.num_real;
	row.appendChild(num_real);

	var diff = document.createElement("td");
	diff.innerHTML = data.diff;
	row.appendChild(diff);
	return row;
}

// 发送整个列表给后台，并下载excel
function getexcel() {
	let httpRequest;
	httpRequest = new XMLHttpRequest();
	let csrftoken = getCookie('csrftoken');

	if (!httpRequest) {
		alert('Giving up :( Cannot create an XMLHTTP instance');
		return false;
	}

	//设置处理响应函数
	httpRequest.onreadystatechange = makeResponse;
	url = "/generatorexcel/download/";
	httpRequest.open('POST', url, true);
	//跨域设置
	httpRequest.setRequestHeader("X-CSRFToken", csrftoken);
	//设置content-type请求头，方便后台识别请求的类型

	// 获取数据发送数据
	exceldata = getexceldata()
	httpRequest.send(JSON.stringify(exceldata));
	function makeResponse() {
		if (httpRequest.readyState == 4 && httpRequest.status == 200){
			var disp = httpRequest.getResponseHeader('Content-Disposition');
			console.log(disp)
            if (disp && disp.search('attachment') != -1) { //判断是否为文件
                var form = $('<form action="' +  url + '" method="post"></form>');
                $('body').append(form);
                form.submit(); //自动提交
                console.log(1);
            }
		}
	}
}

// 获取列表的数据
function getexceldata() {
	var data = [];
	for (var i = 1; i < table.rows.length; i++) {
		var value = {
			"num_require": table.rows[i].cells[1].innerHTML,
			"num_real": table.rows[i].cells[2].innerHTML,
			"diff": table.rows[i].cells[3].innerHTML
		};
		data.push(value);
	}
	return data;
}

//从cookie获取信息
function getCookie(name) {
    var cookieValue = null;
    if (document.cookie && document.cookie !== '') {
        var cookies = document.cookie.split(';');
        for (var i = 0; i < cookies.length; i++) {
            var cookie = cookies[i].trim();
            // Does this cookie string begin with the name we want?
            if (cookie.substring(0, name.length + 1) === (name + '=')) {
                cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
                break;
            }
        }
    }
    return cookieValue;
}