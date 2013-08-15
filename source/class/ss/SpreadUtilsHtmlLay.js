qx.Class.define("ss.SpreadUtilsHtmlLay",
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
     * @param node {Node} TODOC
     * @return {Map} TODOC
     */
    getBorderBoxWidth : function( /* HTMLElement */ node)
    {
      //	summary
      //	Returns the dimensions of the passed element based on border-box sizing.
      node = ss.SpreadUtilsHtmlDom.byId(node);

      // OSY return { width: node.offsetWidth, height: node.offsetHeight };	//	object
      return { width : node.offsetWidth };  //	object
    },

    // return node.offsetWidth;
    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @return {Map} TODOC
     */
    getBorderBoxHeight : function( /* HTMLElement */ node)
    {
      //	summary
      //	Returns the dimensions of the passed element based on border-box sizing.
      node = ss.SpreadUtilsHtmlDom.byId(node);

      // OSY return { width: node.offsetWidth, height: node.offsetHeight };	//	object
      return { height : node.offsetHeight };  //	object
    }

    // return node.offsetHeight;

  }
});
