<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="PageTab.ascx.vb" Inherits="MBSC.UICtl.PageTab" %>

<style>
    #header-inner 
    {
        background-position: center;
        margin-left: auto;
        margin-right: auto;
        height:155px;
    }

    #outer-wrapper {
        width: 960px;
        margin: 0px auto 0px;
        padding: 0px;
        text-align: left;
        height:155px;
    }
</style>

<div id="header-inner" >
    <div style="background: transparent">
        <div style="height: 70px;"></div>
        <div id="outer-wrapper" style="background: transparent; border-width: 0px; font-size:14pt;font-weight:bold;display:none">
            <a href="http://mbscorg.blogspot.tw/">請連結 MBSC 佛陀原始正法學會官網http://mbscorg.blogspot.tw </a>
        </div>
    </div>
</div>

<!--Menu dropdown include-->
 <nav class="navbar navbar-default navbar-fixed-top" role="navigation">
                <div class="container">
                    <!-- Brand and toggle get grouped for better mobile display -->
                    <div class="navbar-header page-scroll">
                    
                        <button type="button" class="navbar-toggle" style="background-color:#F9BC83" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1">                            
                           <B>請登入</B>
                        </button>
                        <a class="navbar-brand page-scroll" href="#page-top"><img class="img-responsive" src="/img/logotext.png" alt="Responsive image"/></a>
                    </div>
                    <!-- Collect the nav links, forms, and other content for toggling -->
                    <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
                         <!-- #include virtual="~/inc/vTab.inc" -->
                    </div>
         <!-- /.navbar-collapse -->
     </div>
     <!-- /.container -->
 </nav>
<!--/.Menu dropdown include-->

<script type="text/javascript">
    jQuery(document).ready(function () {
        var offset = 220;
        var duration = 500;
        jQuery(window).scroll(function () {
            if (jQuery(this).scrollTop() > offset) {
                jQuery('.back-to-top').fadeIn(duration);
            } else {
                jQuery('.back-to-top').fadeOut(duration);
            }            
        });

        jQuery('.back-to-top').click(function (event) {
            event.preventDefault();
            jQuery('html, body').animate({ scrollTop: 0 }, duration);
            return false;
        })
    });
</script>
