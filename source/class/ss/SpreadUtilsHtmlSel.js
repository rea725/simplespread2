qx.Class.define("qx.ui.spread.SpreadUtilsHtmlSel",
{
  /*
  *****************************************************************************
     STATICS
  *****************************************************************************
  */

  statics :
  {
    /**
     * TODOC
     *
     * @type static
     * @param element {var} TODOC
     * @return {boolean} TODOC
     */
    disableSelection : function( /* DomNode */ element)
    {
      // summary: disable selection on a node
      element = qx.ui.spread.SpreadUtilsHtmlDom.byId(element) || document.getElementsByTagName("body")[0];

      var client = qx.core.Client.getInstance();

      if (client.isGecko()) {
        element.style.MozUserSelect = "none";
      } else if (client.isSafari2()) {
        element.style.KhtmlUserSelect = "none";
      } else if (client.isMshtml()) {
        element.unselectable = "on";
      } else {
        return false;
      }

      return true;
    },


    /**
     * TODOC
     *
     * @type static
     * @param element {var} TODOC
     * @return {boolean} TODOC
     */
    enableSelection : function( /* DomNode */ element)
    {
      // summary: enable selection on a node
      element = qx.ui.spread.SpreadUtilsHtmlDom.byId(element) || document.getElementsByTagName("body")[0];

      var client = qx.core.Client.getInstance();

      if (client.isGecko()) {
        element.style.MozUserSelect = "";
      } else if (client.isSafari2()) {
        element.style.KhtmlUserSelect = "";
      } else if (client.isMshtml()) {
        element.unselectable = "off";
      } else {
        return false;
      }

      return true;
    }
  }
});
