qx.Class.define("qx.ui.spread.SpreadUtilsHtmlCommon",
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
     * @param type {var} TODOC
     * @return {var} TODOC
     */
    getParentByType : function( /* HTMLElement */ node, /* string */ type)
    {
      //	summary
      //	Returns the first ancestor of node with tagName type.
      var _document = this.document;
      var parent = qx.ui.spread.SpreadUtilsHtmlDom.byId(node);
      type = type.toLowerCase();

      while ((parent) && (parent.nodeName.toLowerCase() != type))
      {
        // if(parent==(_document["body"]||_document["documentElement"])){
        // return null;
        // }
        parent = parent.parentNode;
      }

      // alert("Node " + node + " type" + type);
      return parent;  //	HTMLElement
    },


    /**
     * TODOC
     *
     * @type static
     * @param e {Event} TODOC
     * @return {var} TODOC
     */
    getCursorPosition : function( /* DOMEvent */ e)
    {
      //	summary
      //	Returns the mouse position relative to the document (not the viewport).
      //	For example, if you have a document that is 10000px tall,
      //	but your browser window is only 100px tall,
      //	if you scroll to the bottom of the document and call this function it
      //	will return {x: 0, y: 10000}

      var cursor =
      {
        x : 0,
        y : 0
      };

      if (e.getPageX() || e.getPageY())
      {
        cursor.x = e.getPageX();
        cursor.y = e.getPageY();
      }
      else
      {
        var de = document.documentElement;

        var db = document.body;
        cursor.x = e.getClientX() + ((de || db)["scrollLeft"]) - ((de || db)["clientLeft"]);
        cursor.y = e.getClientY() + ((de || db)["scrollTop"]) - ((de || db)["clientTop"]);
      }

      return cursor;  //	object
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @param attr {var} TODOC
     * @return {null | var} TODOC
     */
    getAttribute : function( /* HTMLElement */ node, /* string */ attr)
    {
      //	summary
      //	Returns the value of attribute attr from node.
      node = qx.ui.spread.SpreadUtilsHtmlDom.byId(node);

      // FIXME: need to add support for attr-specific accessors
      if ((!node) || (!node.getAttribute))
      {
        // if(attr !== 'nwType'){
        //	alert("getAttr of '" + attr + "' with bad node");
        // }
        return null;
      }

      var ta = typeof attr == 'string' ? attr : new String(attr);

      // first try the approach most likely to succeed
      var v = node.getAttribute(ta.toUpperCase());

      if ((v) && (typeof v == 'string') && (v != "")) {
        return v;  //	string
      }

      // try returning the attributes value, if we couldn't get it as a string
      if (v && v.value) {
        return v.value;  //	string
      }

      // this should work on Opera 7, but it's a little on the crashy side
      if ((node.getAttributeNode) && (node.getAttributeNode(ta))) {
        return (node.getAttributeNode(ta)).value;  //	string
      } else if (node.getAttribute(ta)) {
        return node.getAttribute(ta);  //	string
      } else if (node.getAttribute(ta.toLowerCase())) {
        return node.getAttribute(ta.toLowerCase());  //	string
      }

      return null;  //	string
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @param attr {var} TODOC
     * @return {var} TODOC
     */
    hasAttribute : function( /* HTMLElement */ node, /* string */ attr)
    {
      //	summary
      //	Determines whether or not the specified node carries a value for the attribute in question.
      return this.getAttribute(qx.ui.spread.SpreadUtilsHtmlDom.byId(node), attr) ? true : false;  //	boolean
    }
  }
});
