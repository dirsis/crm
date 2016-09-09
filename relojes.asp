<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Buffer=True%> 
<!-- #include file="odbc.asp" -->
<!DOCTYPE html>
<!--[if IE 8]> <html lang="en" class="ie8"> <![endif]-->
<!--[if IE 9]> <html lang="en" class="ie9"> <![endif]-->
<!--[if !IE]><!--> <html lang="en"> <!--<![endif]-->
<head>
	<title>Charts &amp; Countdowns | Unify - Responsive Website Template</title>
	<!-- Meta -->
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<meta name="description" content="">
	<meta name="author" content="">

	<!-- Favicon -->
	<link rel="shortcut icon" href="favicon.ico">

	<!-- Web Fonts -->
	<link rel="stylesheet" href="http://fonts.googleapis.com/css?family=Open+Sans:400,300,600&amp;subset=cyrillic,latin">

	<!-- CSS Global Compulsory -->
	<link rel="stylesheet" href="PRUEBA/bootstrap.min.css">
	<link rel="stylesheet" href="PRUEBA/style.css">

	<!-- CSS Header and Footer -->
	<link rel="stylesheet" href="PRUEBA/headers/header-default.css">
	<link rel="stylesheet" href="PRUEBA/footers/footer-v1.css">

	<!-- CSS Implementing Plugins -->
	<link rel="stylesheet" href="PRUEBA/animate.css">
	<link rel="stylesheet" href="PRUEBA/line-icons.css">
	<link rel="stylesheet" href="PRUEBA/font-awesome.min.css">

	<!-- CSS Theme -->
	<link rel="stylesheet" href="PRUEBA/default.css" id="style_color">
	<link rel="stylesheet" href="PRUEBA/dark.css">

	<!-- CSS Customization -->
	<link rel="stylesheet" href="PRUEBA/custom.css">
</head>

<body>
	<div class="wrapper">

		<!--=== Content Part ===-->
		<div class="container content">
			<div class="row">
				<!-- Begin Sidebar Menu -->
				<!-- End Sidebar Menu -->
                <!-- Begin Content -->
				<div class="col-md-9">
					<div class="tag-box tag-box-v3">
						<div class="headline"><h2>Estadisticas</h2></div>
						<p>Los valores son actualizados al momento de enfocar la pantalla.</p><br><br>
						<!-- Pie Charts v1 -->
						<div class="row pie-progress-charts margin-bottom-60">
							<div class="inner-pchart col-md-4">
								<div class="circle" id="circle-1"></div>
								<h3 class="circle-title">Gestion Exitosas</h3>
								<p>Se contabilizan las terminadas en exito!</p>
							</div>
							<div class="inner-pchart col-md-4">
								<div class="circle" id="circle-2"></div>
								<h3 class="circle-title">Gestones Caidas</h3>
								<p>Se contabilizan las sin retorno</p>
							</div>
							<div class="inner-pchart col-md-4">
								<div class="circle" id="circle-3"></div>
								<h3 class="circle-title">Gestiones en Lucha</h3>
								<p>Las que tienen esperanza</p>
							</div>
						</div>
						<div class="margin-bottom-60"><hr></div>
						<div class="row pie-progress-charts margin-bottom-60">
							<div class="inner-pchart col-md-3">
								<div class="circle" id="circles-1"></div>
								<h3 class="circle-title">Clientes Concretados </h3>
							</div>
							<div class="inner-pchart col-md-3">
								<div class="circle" id="circles-2"></div>
								<h3 class="circle-title">Clientes En Charla </h3>
							</div>
							<div class="inner-pchart col-md-3">
								<div class="circle" id="circles-3"></div>
								<h3 class="circle-title">Cliente Perdidos </h3>
							</div>
							<div class="inner-pchart col-md-3">
								<div class="circle" id="circles-4"></div>
								<h3 class="circle-title">Cliente Deseados </h3>
							</div>
						</div>
					</div>
			  </div>
				<!-- End Content -->
		  </div>
		</div><!--/container-->
	</div><!--/End Wrapepr-->
	<i class="style-switcher-btn fa fa-cogs hidden-xs"></i>
<script type="text/javascript" src="PRUEBA/jquery.min.js"></script>
	<script type="text/javascript" src="PRUEBA/jquery-migrate.min.js"></script>
	<script type="text/javascript" src="PRUEBA/bootstrap.min.js"></script>
	<!-- JS Implementing Plugins -->
	<script type="text/javascript" src="PRUEBA/back-to-top.js"></script>
	<script type="text/javascript" src="PRUEBA/smoothScroll.js"></script>
	<script type="text/javascript" src="PRUEBA/waypoints.min.js"></script>
	<script type="text/javascript" src="PRUEBA/jquery.counterup.min.js"></script>
	<script type="text/javascript" src="PRUEBA/circles.js"></script>
	<!-- JS Customization -->
	<script type="text/javascript" src="PRUEBA/custom.js"></script>
	<!-- JS Page Level -->
	<script type="text/javascript" src="PRUEBA/app.js"></script>
	<script type="text/javascript" src="PRUEBA/style-switcher.js"></script>
	<%
	sql="select count(idclien) as cantidad from clien"
	set rs = oConn.Execute(sql)
	xclien=rs("cantidad")
	sql="select count(idgestion) as cantidad from gestion where estado=0"
	set rs = oConn.Execute(sql)
	xgestion0=rs("cantidad")
	sql="select count(idgestion) as cantidad from gestion where estado=1"
	set rs = oConn.Execute(sql)
	xgestion1=rs("cantidad")
	sql="select count(idgestion) as cantidad from gestion where estado=2"
	set rs = oConn.Execute(sql)
	xgestion2=rs("cantidad")
	
	xgestiont=xgestion0+xgestion1+xgestion2
	xgestion0=xgestion0*100/xgestiont
	xgestion1=xgestion1*100/xgestiont
	xgestion2=xgestion2*100/xgestiont
	
	%>
	<script>
var CirclesMaster = function () {

    return {

        //Circles Master v1
        initCirclesMaster1: function () {
        	//Circles 1
		    Circles.create({
		        id:         'circle-1',
		        percentage: <%=xgestion0%>,
		        radius:     80,
		        width:      8,
		        number:     <%=xgestion0%>,
		        text:       '%',
		        colors:     ['#eee', '#72c02c'],
		        duration:   2000
		    })

        	//Circles 2
		    Circles.create({
		        id:         'circle-2',
		        percentage: <%=xgestion1%>,
		        radius:     80,
		        width:      8,
		        number:     <%=xgestion1%>,
		        text:       '%',
		        colors:     ['#eee', '#72c02c'],
		        duration:   2000
		    })

        	//Circles 3
		    Circles.create({
		        id:         'circle-3',
		        percentage: <%=xgestion2%>,
		        radius:     80,
		        width:      8,
		        number:     <%=xgestion2%>,
		        text:       '%',
		        colors:     ['#eee', '#72c02c'],
		        duration:   2000
		    })



		    //Circles 5
		    Circles.create({
		        id:         'circle-5',
		        percentage: 20,
		        radius:     35,
		        width:      2,
		        number:     20,
		        text:       '%',
		        colors:     ['#eee', '#9B6BCC'],
		        duration:   2000
		    })

		    //Circles 6
		    Circles.create({
		        id:         'circle-6',
		        percentage: 30,
		        radius:     80,
		        width:      3,
		        number:     30,
		        text:       '%',
		        colors:     ['#eee', '#72c02c'],
		        duration:   2000
		    })

        	//Circles 7
		    Circles.create({
		        id:         'circle-7',
		        percentage: 40,
		        radius:     80,
		        width:      3,
		        number:     40,
		        text:       '%',
		        colors:     ['#eee', '#72c02c'],
		        duration:   2000
		    })

        	//Circles 8
		    Circles.create({
		        id:         'circle-8',
		        percentage: 50,
		        radius:     80,
		        width:      3,
		        number:     50,
		        text:       '%',
		        colors:     ['#eee', '#72c02c'],
		        duration:   2000
		    })

		    //Circles 9
		    Circles.create({
		        id:         'circle-9',
		        percentage: 60,
		        radius:     80,
		        width:      3,
		        number:     60,
		        text:       '%',
		        colors:     ['#eee', '#72c02c'],
		        duration:   2000
		    })
        },
        
        //Circles Master v2
        initCirclesMaster2: function () {
		    var colors = [
		        ['#D3B6C6', '#9B6BCC'], ['#C9FF97', '#72c02c'], ['#BEE3F7', '#3498DB'], ['#FFC2BB', '#E74C3C']
		    ];

		    for (var i = 1; i <= 4; i++) {
		        var child = document.getElementById('circles-' + i),
		            percentage = 45 + (i * 9);
		            
		        Circles.create({
		            id:         child.id,
		            percentage: percentage,
		            radius:     70,
		            width:      2,
		            number:     percentage / 1,
		            text:       '%',
		            colors:     colors[i - 1],
		            duration:   2000,
		        });
		    }	    
        }

    };
    
}();
	</script>
	<script type="text/javascript">
		jQuery(document).ready(function() {
			App.init();
			App.initCounter();
			StyleSwitcher.initStyleSwitcher();
			CirclesMaster.initCirclesMaster1();
			CirclesMaster.initCirclesMaster2();
		});
	</script>
	<!--[if lt IE 9]>
	<script src="assets/plugins/respond.js"></script>
	<script src="assets/plugins/html5shiv.js"></script>
	<script src="assets/plugins/placeholder-IE-fixes.js"></script>
	<![endif]-->
</body>
</html>
