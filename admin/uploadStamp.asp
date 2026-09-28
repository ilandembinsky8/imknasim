<html>
    <head>
        <title></title>
        <meta charset="UTF-8">
        <style>
            .custom-file-input::-webkit-file-upload-button {
              visibility: hidden;
            }
            .custom-file-input::before {
                content: 'בחר קובץ';
                display: inline-block;
                width: 200px;
                height: 40px;
                background-color: #EAEAEA;
                box-sizing: border-box;
                border-radius: 8px;
                border: 0px;
                margin-top: 20px;
                font-size: 24px;
                text-align:center;
                padding-top:5px;
            }
            .custom-file-input:hover::before {
              border-color: black;
            }
            .custom-file-input:active::before {
              background: -webkit-linear-gradient(top, #e3e3e3, #f9f9f9);
            }
        </style>
    </head>
	<body style="font-family:Arial; font-size:24px;direction:rtl;text-align:right">
			<FORM NAME="MyForm" dir=ltr METHOD="POST" ENCTYPE="multipart/form-data" ACTION="uploadStamp1.asp" ID="Form1">
		    <div style="float:right;width:100%">
                <div style="margin:0px auto;width:440px;height:200px;background-color:#DADADA;box-sizing:border-box;padding:20px">
                <div style="float:right;width:100%">קובץ / תמונה</div> 
                <div style="float:right;width:100%;text-align:right"><input class="custom-file-input" type="file" NAME="FILE<%=i%>" ID="File4" value="בחר קובץ"></div>
                <div style="float:right;width:80%;margin-right:50px;margin-top:-10px"><span id="id4button"><input style="width:200px;height:40px;background-color:#EAEAEA;box-sizing:border-box;border-radius:8px;border:0px;margin-top:20px;font-size:24px" type="submit" value="טען קובץ" ID="Button1" onclick="id4button.innerHTML = 'File upload takes time. Please wait;'; MyForm.submit()"></span></div>
				<input name="filepath1" value="<%=filepath1%>" type="hidden">
				<input name="filepath2" value="<%=filepath2%>" type="hidden">
				<input name="fieldname" value="<%=request("id")%>" type="hidden">
                <input name="thefolder" value="<%=request("theFolder")%>" type="hidden" />
				
                </div>
            </div>
			</form>
	</body>
</html>
