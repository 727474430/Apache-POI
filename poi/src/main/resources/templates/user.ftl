<#assign base = request.contextPath />
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="author" content="wangliang">
    <base id="base" href="${base}">
    <title>Spring Boot - Freemarker</title>
    <!-- Bootstrap core CSS -->
    <link href="//cdn.jsdelivr.net/bootstrap/3.3.6/css/bootstrap.min.css" rel="stylesheet">
    <script src="http://libs.baidu.com/jquery/2.0.0/jquery.min.js"></script>
    <script src="http://libs.baidu.com/bootstrap/3.3.6/js/bootstrap.min.js"></script>
</head>

<body>
<div class="container">
    <div class="page-header">
        <h1 align="center"><span style="color: greenyellow;">User Lkist</span></h1>
    </div>
    <div>
        <form id="userSub" action="#" method="post" enctype="multipart/form-data">
            <table class="table table-striped">
                <caption>
                    现有用户列表
                    <button id="exportExcel" class="btn btn-primary" style="float: right">
                        导出Excel
                    </button>
                    <button id="importExcel" class="btn btn-primary" style="float: right">
                        导入Excel
                    </button>
                    <button class="btn btn-primary" style="float: right">
                        <input type="file" name="file" size="40">
                    </button>
                </caption>
                <thead>
                <tr>
                    <th>姓名</th>
                    <th>年龄</th>
                    <th>性别</th>
                    <th>邮箱</th>
                    <th>手机</th>
                </tr>
                </thead>
                <tbody>
                <#if users??>
                    <#list users as user>
                    <tr>
                        <td>${user.name}</td>
                        <td>${user.age}</td>
                        <td>${user.sex}</td>
                        <td>${user.email}</td>
                        <td>${user.phone}</td>
                    </tr>
                    </#list>
                </#if>
                </tbody>
            </table>
        </form>
    </div>
</div>

<div class="footer">
    <div class="container">
        <p class="text-muted" style="float: right;">©2017 WangLiang</p>
    </div>
</div>
</body>
<script type="text/javascript">
    $(function () {
        $("#exportExcel").click(function () {
            $("form[id=userSub]").attr('action', '${base}/users/v1/export');
        });

        $("#importExcel").click(function () {
            $("form[id=userSub]").attr('action', '${base}/users/v1/import');
        });

        $("#ajaxBtn").click(function () {
            var formData = $("#ajaxForm").serializeObject();
            var pageData = $("#pageForm").serializeObject();
            var data = $.extend({}, pageData, formData);
            $.ajax({
                url: "${base}/json/submit",
                type: "post",
                contentType: "application/json;charset=UTF-8",
                dataType: "json",
                data: JSON.stringify(data),
                success: function (data) {
                    console.log(data);
                }
            })
        });
    })


    $.fn.serializeObject = function () {
        var result = {};
        var formData = this.serializeArray();
        $.each(formData, function (i, v) {
            result[v.name] = v.value;
        })
        return result;
    };
</script>

</html>