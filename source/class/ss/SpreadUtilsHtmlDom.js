


qx.Class.define("ss.SpreadUtilsHtmlDom",
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
     * @param ref {var} TODOC
     * @param force {Boolean} TODOC
     * @return {boolean} TODOC
     */
    insertBefore : function( /* Node */ node, /* Node */ ref, /* boolean? */ force)
    {
      //	summary:
      //		Try to insert node before ref
      if ((force != true) && (node === ref || node.nextSibling === ref)) {
        return false;
      }

      var parent = ref.parentNode;
      parent.insertBefore(node, ref);
      return true;  //	boolean
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @param ref {var} TODOC
     * @param force {Boolean} TODOC
     * @return {boolean | var} TODOC
     */
    insertAfter : function( /* Node */ node, /* Node */ ref, /* boolean? */ force)
    {
      //	summary:
      //		Try to insert node after ref
      var pn = ref.parentNode;

      if (ref == pn.lastChild)
      {
        if ((force != true) && (node === ref)) {
          return false;  //	boolean
        }

        pn.appendChild(node);
      }
      else
      {
        return this.insertBefore(node, ref.nextSibling, force);  //	boolean
      }

      return true;  //	boolean
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @param ref {var} TODOC
     * @param position {var} TODOC
     * @return {boolean | var} TODOC
     */
    insertAtPosition : function( /* Node */ node, /* Node */ ref, /* string */ position)
    {
      //	summary:
      //		attempt to insert node in relation to ref based on position
      if ((!node) || (!ref) || (!position)) {
        return false;  //	boolean
      }

      switch(position.toLowerCase())
      {
        case "before":
          return qx.Proto.insertBefore(node, ref);  //	boolean

        case "after":
          return qx.Proto.insertAfter(node, ref);  //	boolean

        case "first":
          if (ref.firstChild) {
            return qx.Proto.insertBefore(node, ref.firstChild);  //	boolean
          }
          else
          {
            ref.appendChild(node);
            return true;  //	boolean
          }

          break;

        default:
            // aka: last
          ref.appendChild(node);
          return true;  //	boolean
      }
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @param containingNode {var} TODOC
     * @param insertionIndex {var} TODOC
     * @return {boolean | var} TODOC
     */
    insertAtIndex : function( /* Node */ node, /* Element */ containingNode, /* number */ insertionIndex)
    {
      //	summary:
      //		insert node into child nodes nodelist of containingNode at
      //		insertionIndex. insertionIndex should be between 0 and
      //		the number of the childNodes in containingNode. insertionIndex
      //		specifys after how many childNodes in containingNode the node
      //		shall be inserted. If 0 is given, node will be appended to
      //		containingNode.
      var siblingNodes = containingNode.childNodes;

      // if there aren't any kids yet, just add it to the beginning
      if (!siblingNodes.length || siblingNodes.length == insertionIndex)
      {
        containingNode.appendChild(node);
        return true;  //	boolean
      }

      if (insertionIndex == 0) {
        return qx.Proto.prependChild(node, containingNode);  //	boolean
      }

      // otherwise we need to walk the childNodes
      // and find our spot
      return qx.Proto.insertAfter(node, siblingNodes[insertionIndex - 1]);  //	boolean
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @param text {var} TODOC
     * @return {var} TODOC
     */
    textContent : function( /* Node */ node, /* string */ text)
    {
      //	summary:
      //		implementation of the DOM Level 3 attribute; scan node for text
      if (arguments.length > 1)
      {
        var _document = document;
        qx.Proto.replaceChildren(node, _document.createTextNode(text));
        return text;  //	string
      }
      else
      {
        if (node.textContent != undefined)
        {  // FF 1.5
          return node.textContent;  //	string
        }

        var _result = "";

        if (node == null) {
          return _result;
        }

        for (var i=0; i<node.childNodes.length; i++)
        {
          switch(node.childNodes[i].nodeType)
          {
            case 1:  // ELEMENT_NODE
            case 5:  // ENTITY_REFERENCE_NODE
              _result += qx.Proto.textContent(node.childNodes[i]);
              break;

            case 3:  // TEXT_NODE
            case 2:  // ATTRIBUTE_NODE
            case 4:  // CDATA_SECTION_NODE
              _result += node.childNodes[i].nodeValue;
              break;

            default:
              break;
          }
        }

        return _result;  //	string
      }
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @return {var} TODOC
     */
    hasParent : function( /* Node */ node)
    {
      //	summary:
      //		returns whether or not node is a child of another node.
      return Boolean(node && node.parentNode && qx.Proto.isNode(node.parentNode));  //	boolean
    },


    /**
     * TODOC
     *
     * @type static
     * @param id {var} TODOC
     * @param doc {Document} TODOC
     * @return {var} TODOC
     */
    byId : function( /* String */ id, /* DocumentElement */ doc)
    {
      // summary:
      // 		similar to other library's "$" function, takes a string
      // 		representing a DOM id or a DomNode and returns the
      // 		corresponding DomNode. If a Node is passed, this function is a
      // 		no-op. Returns a single DOM node or null, working around
      // 		several browser-specific bugs to do so.
      // id: DOM id or DOM Node
      // doc:
      //		optional, defaults to the current value of dj_currentDocument.
      //		Can be used to retreive node references from other documents.
      if ((id) && ((typeof id == "string") || (id instanceof String)))
      {
        if (!doc) {
          doc = dj_currentDocument;
        }

        var ele = doc.getElementById(id);

        // workaround bug in IE and Opera 8.2 where getElementById returns wrong element
        if (ele && (ele.id != id) && doc.all)
        {
          ele = null;

          // get all matching elements with this id
          eles = doc.all[id];

          if (eles)
          {
            // if more than 1, choose first with the correct id
            if (eles.length)
            {
              for (var i=0; i<eles.length; i++)
              {
                if (eles[i].id == id)
                {
                  ele = eles[i];
                  break;
                }
              }
            }

            // return 1 and only element
            else
            {
              ele = eles;
            }
          }
        }

        return ele;  // DomNode
      }

      return id;  // DomNode
    },


    /**
     * TODOC
     *
     * @type static
     * @param node {Node} TODOC
     * @param ancestor {var} TODOC
     * @param guaranteeDescendant {var} TODOC
     * @return {boolean} TODOC
     */
    isDescendantOf : function( /* Node */ node, /* Node */ ancestor, /* boolean? */ guaranteeDescendant)
    {
      //	summary
      //	Returns boolean if node is a descendant of ancestor
      // guaranteeDescendant allows us to be a "true" isDescendantOf function
      if (guaranteeDescendant && node) {
        node = node.parentNode;
      }

      while (node)
      {
        if (node == ancestor) {
          return true;  //	boolean
        }

        node = node.parentNode;
      }

      return false;  //	boolean
    }
  }
});

this.ELEMENT_NODE = 1;
this.ATTRIBUTE_NODE = 2;
this.TEXT_NODE = 3;
this.CDATA_SECTION_NODE = 4;
this.ENTITY_REFERENCE_NODE = 5;
this.ENTITY_NODE = 6;
this.PROCESSING_INSTRUCTION_NODE = 7;
this.COMMENT_NODE = 8;
this.DOCUMENT_NODE = 9;
this.DOCUMENT_TYPE_NODE = 10;
this.DOCUMENT_FRAGMENT_NODE = 11;
this.NOTATION_NODE = 12;
