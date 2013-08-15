qx.Class.define("ss.SpreadUtilsHtmlStyle",
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
     * @param classStr {var} TODOC
     * @return {boolean | var} TODOC
     */
    addClass : function( /* HTMLElement */ node, /* string */ classStr)
    {
      //	summary
      //	Adds the specified class to the end of the class list on the
      //	passed &node;. Returns &true; or &false; indicating success or failure.
      if (this.hasClass(node, classStr)) {
        return false;
      }

      classStr = (this.getClass(node) + " " + classStr).replace(/^\s+|\s+$/g, "");
      return this.setClass(node, classStr);  //	boolean
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @param classStr {var} TODOC
     * @return {boolean} TODOC
     */
    setClass : function( /* HTMLElement */ node, /* string */ classStr)
    {
      //	summary
      //	Clobbers the existing list of classes for the node, replacing it with
      //	the list given in the 2nd argument. Returns true or false
      //	indicating success or failure.
      node = ss.SpreadUtilsHtmlDom.byId(node);
      var cs = new String(classStr);

      try
      {
        if (typeof node.className == "string") {
          node.className = cs;
        }
        else if (node.setAttribute)
        {
          node.setAttribute("class", classStr);
          node.className = cs;
        }
        else
        {
          return false;
        }
      }
      catch(e)
      {
        this.debug("setClass() failed", e);
      }

      return true;
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @param classStr {var} TODOC
     * @param allowPartialMatches {var} TODOC
     * @return {boolean} TODOC
     */
    removeClass : function( /* HTMLElement */ node, /* string */ classStr, /* boolean? */ allowPartialMatches)
    {
      //	summary
      //	Removes the className from the node;. Returns true or false indicating success or failure.
      try
      {
        if (!allowPartialMatches) {
          var newcs = this.getClass(node).replace(new RegExp('(^|\\s+)' + classStr + '(\\s+|$)'), "$1$2");
        } else {
          var newcs = this.getClass(node).replace(classStr, '');
        }

        this.setClass(node, newcs);
      }
      catch(e)
      {
        this.debug("removeClass() failed", e);
      }

      return true;  //	boolean
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @return {string | var} TODOC
     */
    getClass : function( /* HTMLElement */ node)
    {
      //	summary
      //	Returns the string value of the list of CSS classes currently assigned directly
      //	to the node in question. Returns an empty string if no class attribute is found;
      node = ss.SpreadUtilsHtmlDom.byId(node);

      if (!node) {
        return "";
      }

      var cs = "";

      if (node.className) {
        cs = node.className;
      } else if (ss.SpreadUtilsHtmlCommon.hasAttribute(node, "class")) {
        cs = ss.SpreadUtilsHtmlCommon.getAttribute(node, "class");
      }

      return cs.replace(/^\s+|\s+$/g, "");  //	string
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @return {var} TODOC
     */
    getClasses : function( /* HTMLElement */ node)
    {
      //	summary
      //	Returns an array of CSS classes currently assigned directly to the node in question.
      //	Returns an empty array if no classes are found;
      var c = this.getClass(node);
      return (c == "") ? [] : c.split(/\s+/g);  //	array
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @param classname {var} TODOC
     * @return {var} TODOC
     */
    hasClass : function( /* HTMLElement */ node, /* string */ classname)
    {
      //	summary
      //	Returns whether or not the specified classname is a portion of the
      //	class list currently applied to the node. Does not cover cascaded
      //	styles, only classes directly applied to the node.
      return (new RegExp('(^|\\s+)' + classname + '(\\s+|$)')).test(this.getClass(node))  //	boolean;
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @param cssSelector {var} TODOC
     * @param inValue {var} TODOC
     * @return {var} TODOC
     */
    getComputedStyle : function( /* HTMLElement */ node, /* string */ cssSelector, /* integer? */ inValue)
    {
      //	summary
      //	Returns the computed style of cssSelector on node.
      node = ss.SpreadUtilsHtmlDom.byId(node);

      // cssSelector may actually be in camel case, so force selector version
      var cssSelector = this.toSelectorCase(cssSelector);
      var property = this.toCamelCase(cssSelector);

      if (!node || !node.style) {
        return inValue;
      }
      else if (document.defaultView && ss.SpreadUtilsHtmlDom.isDescendantOf(node, node.ownerDocument))
      {  // W3, gecko, KHTML
        try
        {
          // mozilla segfaults when margin-* and node is removed from doc
          // FIXME: need to figure out a if there is quicker workaround
          var cs = document.defaultView.getComputedStyle(node, "");

          if (cs) {
            return cs.getPropertyValue(cssSelector);  //	integer
          }
        }
        catch(e)
        {  // reports are that Safari can throw an exception above
          if (node.style.getPropertyValue)
          {  // W3
            return node.style.getPropertyValue(cssSelector);  //	integer
          }
          else
          {
            return inValue;  //	integer
          }
        }
      }
      else if (node.currentStyle)
      {  // IE
        return node.currentStyle[property];  //	integer
      }

      if (node.style.getPropertyValue)
      {  // W3
        return node.style.getPropertyValue(cssSelector);  //	integer
      }
      else
      {
        return inValue;  //	integer
      }
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @param cssSelector {var} TODOC
     * @return {var} TODOC
     */
    getStyleProperty : function( /* HTMLElement */ node, /* string */ cssSelector)
    {
      //	summary
      //	Returns the value of the passed style
      node = ss.SpreadUtilsHtmlDom.byId(node);
      return (node && node.style ? node.style[this.toCamelCase(cssSelector)] : undefined);  //	string
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @param cssSelector {var} TODOC
     * @return {var} TODOC
     */
    getStyle : function( /* HTMLElement */ node, /* string */ cssSelector)
    {
      //	summary
      //	Returns the computed value of the passed style
      var value = this.getStyleProperty(node, cssSelector);
      return (value ? value : this.getComputedStyle(node, cssSelector));  //	string || integer
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @return {var} TODOC
     */
    getBackgroundColor : function( /* HTMLElement */ node)
    {
      //	summary
      //	returns the background color of the passed node as a 32-bit color (RGBA)
      node = ss.SpreadUtilsHtmlDom.byId(node);
      var color;

      do
      {
        color = this.getStyle(node, "background-color");

        // Safari doesn't say "transparent"
        if (color.toLowerCase() == "rgba(0, 0, 0, 0)") {
          color = "transparent";
        }

        if (node == document.getElementsByTagName("body")[0])
        {
          node = null;
          break;
        }

        node = node.parentNode;
      }
      while (node && qx.lang.Array.contains([ "transparent", "" ], color));

      if (color == "transparent") {
        color = [ 255, 255, 255, 0 ];
      } else {
        color = this.extractRGB(color);
      }

      return color;  //	array
    },

    // get RGB array from css-style color declarations
    /**
     * TODOC
     *
     * @type static
     * @param color {var} TODOC
     * @return {var} TODOC
     */
    extractRGB : function(color)
    {
      var hex = "0123456789abcdef";
      color = color.toLowerCase();

      if (color.indexOf("rgb") == 0)
      {
        var matches = color.match(/rgba*\((\d+), *(\d+), *(\d+)/i);
        var ret = matches.splice(1, 3);
        return ret;
      }
      else
      {
        var colors = this.hex2rgb(color);

        if (colors) {
          return colors;
        }
        else
        {
          // named color (how many do we support?)
          return this.named[color] || [ 255, 255, 255 ];
        }
      }
    },

    named :
    {
      white  : [ 255, 255, 255 ],
      black  : [ 0, 0, 0 ],
      red    : [ 255, 0, 0 ],
      green  : [ 0, 255, 0 ],
      lime   : [ 0, 255, 0 ],
      blue   : [ 0, 0, 255 ],
      navy   : [ 0, 0, 128 ],
      gray   : [ 128, 128, 128 ],
      silver : [ 192, 192, 192 ]
    },


    /**
     * TODOC
     *
     * @type static
     * @param hex {var} TODOC
     * @return {null | var} TODOC
     */
    hex2rgb : function(hex)
    {
      var hexNum = "0123456789ABCDEF";
      var rgb = new Array(3);

      if (hex.indexOf("#") == 0) {
        hex = hex.substring(1);
      }

      hex = hex.toUpperCase();

      if (hex.replace(new RegExp("[" + hexNum + "]", "g"), "") != "") {
        return null;
      }

      if (hex.length == 3)
      {
        rgb[0] = hex.charAt(0) + hex.charAt(0);
        rgb[1] = hex.charAt(1) + hex.charAt(1);
        rgb[2] = hex.charAt(2) + hex.charAt(2);
      }
      else
      {
        rgb[0] = hex.substring(0, 2);
        rgb[1] = hex.substring(2, 4);
        rgb[2] = hex.substring(4);
      }

      for (var i=0; i<rgb.length; i++) {
        rgb[i] = hexNum.indexOf(rgb[i].charAt(0)) * 16 + hexNum.indexOf(rgb[i].charAt(1));
      }

      return rgb;
    },


    /**
     * TODOC
     *
     * @type static
     * @param r {var} TODOC
     * @param g {var} TODOC
     * @param b {var} TODOC
     * @return {var} TODOC
     */
    rgb2hex : function(r, g, b)
    {
      if (r instanceof Array)
      {
        g = r[1] || 0;
        b = r[2] || 0;
        r = r[0] || 0;
      }

      // TODO OSY
      var ret = Array.map([ r, g, b ], function(x)
      {
        x = new Number(x);
        var s = x.toString(16);

        while (s.length < 2) {
          s = "0" + s;
        }

        return s;
      });

      ret.unshift("#");
      return ret.join("");
    },


    /**
     * TODOC
     *
     * @type static
     * @param selector {var} TODOC
     * @return {var} TODOC
     */
    toSelectorCase : function( /* string */ selector)
    {
      //	summary
      //	Translates a camel cased string to a selector cased one.
      return selector.replace(/([A-Z])/g, "-$1").toLowerCase();  //	string
    },


    /**
     * TODOC
     *
     * @type static
     * @param selector {var} TODOC
     * @return {var} TODOC
     */
    toCamelCase : function( /* string */ selector)
    {
      //	summary
      //	Translates a CSS selector string to a camel-cased one.
      var arr = selector.split('-'), cc = arr[0];

      for (var i=1; i<arr.length; i++) {
        cc += arr[i].charAt(0).toUpperCase() + arr[i].substring(1);
      }

      return cc;  //	string
    }
  }
});
