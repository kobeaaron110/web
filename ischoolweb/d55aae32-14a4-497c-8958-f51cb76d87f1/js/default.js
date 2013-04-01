var _gg = _gg || {};
//_gg.connection = gadget.getContract("mdhs.student");
_gg.connection = gadget.getContract("mdhs.teacher");
//_gg.connection = gadget.getContract("ischool.zzz.student");
var _refGetTeacherWord = {};
var _varClass = "";



jQuery(function () {
	_gg.loadData();
	
    $("#editModal").modal({
        show: false
    });
    $("#editModal").on("hidden", function () {
        $("#editModal #errorMessage").html("");
    });
    $("#editModal").on("show", function () {
        $("#editModal #save-data").button("reset");
    });
	$("#selClassName").change(function () {
		selClassNamechange(this);
	});
    $("#editModal #save-data").click(function () {
        $(this).button("loading");
		//alert("save-data");
		var edit_calss, exam;
		edit_calss = $('#selClassName').val();
		exam = $('#selExam').val();
		if (edit_calss) {
			_gg.SaveWord();
		}
    });
	
	
	$("#editModalChange").click(function () {
        alert("editModalChange");
		var edit_calss = $('#selClassName').val();
		var exam = $('#selExam').val();
		if (edit_calss) {
			_gg.EditWord(edit_calss, exam);
		}
    });
	
	
	/*
    // TODO: 取得學生資訊
    _gg.connection.send({
        service: "_.GetStudent",
        body: '<Request><Condition></Condition></Request>',
        result: function (response, error, http) {
            if (error !== null) {
                $("#mainMsg").html("<div class='alert alert-error'>\n  <button class='close' data-dismiss='alert'>×</button>\n  <strong>呼叫服務失敗或網路異常，請稍候重試!</strong>(GetStudent)\n</div>");
            } else {
                // TODO: 成功
                var ret = '';
                var _ref;
                if (((_ref = response.Response) != null ? _ref.Student : void 0) != null) {
                    $(response.Response.Student).each(function(index, item) {
                        ret = '<p>姓名：' + (item.Name || '') + '</p>' +
                            '<p>學號：' + (item.StudentNumber || '') + '</p>' +
                            '<p>座號：' + (item.SeatNo || '') + '</p>';
                    });
                    $('#container-main').html(ret);
                }
            }
        }
    });
	*/
});
	
	_gg.loadData = function () {
		// TODO: 取得學生資訊
		_gg.connection.send({
			service: "_.GetTeacherWord",
			body: '<Request><Condition></Condition></Request>',
			result: function (response, error, http) {
				if (error !== null) {
					$("#mainMsg").html("<div class='alert alert-error'>\n  <button class='close' data-dismiss='alert'>×</button>\n  <strong>呼叫服務失敗或網路異常，請稍候重試!</strong>(GetStudent)\n</div>");
				} else {
					// TODO: 成功
					var ret = '';
					$("#selClassName option").remove();
					//var _refGetTeacherWord;
					if (((_refGetTeacherWord = response.Response) != null ? _refGetTeacherWord.TeacherWord : void 0) != null) {
						
						$(response.Response.TeacherWord).each(function(index, item) {
							//var content='';
							//for(var key in item){ 
								   //content +="屬性名稱："+ key+" ; 值： "+item[key]+"\n"; 
							//}
							//alert(content);
							//alert($(response.Response.TeacherWord).size());
							
							$("#selClassName").append("<option value='"+item.ClassName+"'"+">"+item.ClassName+"</option>");
							$("#selClassName").val(item.ClassName);
							$("#selExam").val(item.Exam);
							
							ret = '<p>班級名稱：' + (item.ClassName || '') + '</p>' +
								'<p>標題：' + (item.Heading || '') + '</p>' +
								'導師的話：<pre>' + (item.Word || '') + '</pre>' +
								'<p>導師姓名：' + (item.TeacherName || '') + '</p>' +
								'<p>最後更新時間：' + (item.UpdateTime || '') + '</p>';
						});
						//ret = '<p>班級名稱：' + (response.Response.TeacherWord[0].ClassName || '') + '</p>';
						$('#container-main').html(ret);
						selClassNameVal(_varClass);
					}
				}
			}
		});
	};
	
	function selClassNamechange(varThis)
	{
		//alert($(varThis).val());
		alert($('#selClassName').val());
		$(_refGetTeacherWord.TeacherWord).each(function(index, item) {
			if (item.ClassName == $('#selClassName').val())
			{			
				$("#selExam").val(item.Exam);
				var ret = '<p>班級名稱：' + (item.ClassName || '') + '</p>' +
						'<p>標題：' + (item.Heading || '') + '</p>' +
						'導師的話：<pre>' + (item.Word || '') + '</pre>' +
						'<p>導師姓名：' + (item.TeacherName || '') + '</p>' +
						'<p>最後更新時間：' + (item.UpdateTime || '') + '</p>';
				$('#container-main').html(ret);	
			}
		});	
	}
	
	function selClassNameVal(varClass)
	{
		if(varClass == ""){
			return;
		}
		$('#selClassName').val(varClass)
		alert(varClass);
		$(_refGetTeacherWord.TeacherWord).each(function(index, item) {
			if (item.ClassName == varClass)
			{			
				$("#selExam").val(item.Exam);
				var ret = '<p>班級名稱：' + (item.ClassName || '') + '</p>' +
						'<p>標題：' + (item.Heading || '') + '</p>' +
						'導師的話：<pre>' + (item.Word || '') + '</pre>' +
						'<p>導師姓名：' + (item.TeacherName || '') + '</p>' +
						'<p>最後更新時間：' + (item.UpdateTime || '') + '</p>';
				$('#container-main').html(ret);	
			}
		});	
	}
	
	function formatDate(v){  
	  if(typeof v == 'string') v = parseDate(v);  
	  if(v instanceof Date){  
		var y = v.getFullYear();  
		var m = v.getMonth() + 1;  
		var d = v.getDate();  
		var h = v.getHours();  
		var i = v.getMinutes();  
		var s = v.getSeconds();  
		var ms = v.getMilliseconds();     
		if(ms>0) return y + '-' + m + '-' + d + ' ' + h + ':' + i + ':' + s + '.' + ms;  
		if(h>0 || i>0 || s>0) return y + '-' + m + '-' + d + ' ' + h + ':' + i + ':' + s;  
		return y + '-' + m + '-' + d;  
	  }  
	  return '';  
	}  
	
	// TODO: 成績登錄編輯畫面
	_gg.EditWord = function (edit_calss, exam) {
		if (edit_calss) {
			var edit_title, scoreType;
			scoreType = 'PaScore';
			if(exam == 1)
			{
				edit_title = edit_calss + ' 第一次段考';
			}else
			{
				edit_title = edit_calss + ' 第二次段考';
			}
			
			var teacherWords = _refGetTeacherWord.TeacherWord;
			if (teacherWords) {
				var arys = [];
				var ret = "";
				$(teacherWords).each(function(key, value) {
					if(value.ClassName == edit_calss)
					{						
						//alert(value.ClassName);
						$(txtHead).val(value.Heading);
						$(textAreaWord).val(value.Word);
						$(txtName).val(value.TeacherName);
					}		
				});
			}

			//$('#editModal').find('h3').html(edit_title).end().find('fieldset').html(ret.join(''));
			$('#editModal').find('h3').html(edit_title);
			$("#save-data").attr('data-type', scoreType);
			$('#editModal input:text:first').focus();
		}
	};

	// TODO: 儲存導師的話
	_gg.SaveWord = function () {
		//alert("save-data-SaveWord");
		var edit_calss = $('#selClassName').val();
		var exam = $('#selExam').val();
		var heading = $('#txtHead').val();
		var word = $('#textAreaWord').val();
		var teacherName = $('#txtName').val();
		var tmp_Date  = new Date();
		tmp_Date = formatDate(tmp_Date);
		var arys = [];
		arys.push('<TeacherWord>');
		arys.push('<Condition><ClassName>' + (edit_calss || '')  + '</ClassName></Condition>');
		//var tmp_score = $('#'+(value.SCUID)).val();
		//var tmp_score = parseInt(tmp_score, 10) + '';
		
		arys.push('<Field>');
		arys.push('<Exam>' + exam + '</Exam>');
		arys.push('<Heading>' + heading + '</Heading>');
		arys.push('<TeacherName>' + teacherName + '</TeacherName>');
		arys.push('<Word>' + word + '</Word>');
		arys.push('<UpdateTime>' + tmp_Date + '</UpdateTime>');
		arys.push('</Field>');
		arys.push('</TeacherWord>');
		alert(arys.join(''));
		_gg.connection.send({
			service: "_.UpTeacherWord",
			body: '<Request>' + arys.join('') + '</Request>',
			result: function (response, error, http) {
				if (error !== null) {
					_gg.set_error_message('#errorMessage', 'UpTeacherWord', error);
				} else {
					// TODO: 儲存成功
					alert("儲存成功");
					_gg.loadData();
					_varClass = edit_calss;
					$('#editModal').modal("hide");
				}
			}
		});
	};
	
	// TODO: 錯誤訊息
	_gg.set_error_message = function(select_str, serviceName, error) {
		var tmp_msg = '<i class="icon-white icon-info-sign my-err-info"></i><strong>呼叫服務失敗或網路異常，請稍候重試!</strong>(' + serviceName + ')';
		if (error !== null) {
			if (error.dsaError) {
				if (error.dsaError.status === "504") {
					switch (error.dsaError.message) {
						 case '501':
							tmp_msg = '<strong>很抱歉，您無存取資料權限！</strong>';
							break;
						case '502':
							tmp_msg = '<strong>儲存失敗，目前未開放填寫!</strong>';
							break;
					}
				}
			}
			$(select_str).html("<div class='alert alert-error'>\n  <button class='close' data-dismiss='alert'>×</button>\n  " + tmp_msg + "\n</div>");
			$('.my-err-info').click(function(){alert('請拍下此圖，並與客服人員連絡，謝謝您。\n' + JSON.stringify(error, null, 2))});
		}
	};
	
	
