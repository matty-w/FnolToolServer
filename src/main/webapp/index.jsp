<html>
<head>
<style><%@include file="/pageStyle.css"%></style>
</head>
<body>

	<h2>FNOL Report</h2>

	


	<form id="emailUploadForm" action="upload" method="post" enctype="multipart/form-data">
	
	
		
		
		
		<input id="flatFileSearch" type='file' name='file' accept=".csv"/> 
		<input id="claimsFileSearch" type='file' name='file' accept=".xlsx"/> 
		<input id="emailSearch" type='file' name='file' accept=".msg" multiple/> <input id="emailExecute" type='submit'>

	</form>
	
	
	<script>
		function showClaimsForm() {
			document.getElementById("emailUploadForm").style.display = "block";
			document.getElementById("flatFileSearch").style.display = "block";
			document.getElementById("emailSearch").style.display = "none";
			document.getElementById("flatFileExecute").style.display = "block";
			document.getElementById("emailExecute").style.display = "none";
		}
	</script>

	<script>
		function showEmailsForm() {
			document.getElementById("emailUploadForm").style.display = "block";
			document.getElementById("flatFileSearch").style.display = "none";
			document.getElementById("emailSearch").style.display = "block";
			document.getElementById("flatFileExecute").style.display = "none";
			document.getElementById("emailExecute").style.display = "block";
		}
	</script>

</body>
</html>
