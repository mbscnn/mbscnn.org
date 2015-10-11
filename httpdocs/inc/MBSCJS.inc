<Script lang="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/JSUtil.js"%>" ></Script>
<!-- 需畫面列印 請include 『printscript.inc』  -->
<script lang="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/KeyCheck.js"%>" ></script>
<script language="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/json2.js"%>"></script>
<script src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/jquery.js"%>" type="text/javascript"></script>
<script src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/bootstrap.min.js"%>" type="text/javascript"></script>
<script src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/scrolling-nav.js"%>" type="text/javascript"></script>
<script src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/jquery.easing.min.js"%>" type="text/javascript"></script>
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
      jQuery(function($) {
    
    $("table").removeAttr("style");
    $("table").addClass("table table-bordered");
});
    
</script>