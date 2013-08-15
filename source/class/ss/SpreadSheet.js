
/**
 * A SpreadSheet Sheet Object
 * Author : O. Sadey
 */
qx.Class.define("ss.SpreadSheet",
{
  extend : qx.ui.layout.VerticalBoxLayout,




  /*
  *****************************************************************************
     CONSTRUCTOR
  *****************************************************************************
  */

  construct : function(sheetName)
  {
    qx.ui.layout.VerticalBoxLayout.call(this);

    this.auto();

    this.sheetName = sheetName;
    this.formulaField = null;
    this.funcButton = null;
    this.initRowsCount = 100;
    this.initColsCount = 20;

    this.copiedCells = [];

    // Cells dependency object array
    this.depCells = new Array();

    this.DECIMAL_SEP = ",";

    this.CELL_TYPES =
    {
      NUMBER  : 0,
      STRING  : 1,
      DATE    : 2,
      FORMULA : 3
    };

    this.FORMATTING_TYPES =
    {
      FONT      : 0,
      FONT_SIZE : 1,
      COLOR     : 2,
      BG_COLOR  : 3,
      BOLD      : 4,
      ITALIC    : 5,
      UNDERLINE : 6,
      ALIGN     : 7
    };

    this.FORMULAS =
    {
      SUM : 0,
      AVG : 0
    };

    // serif', 'sans-serif', 'cursive', 'fantasy', and 'monospace
    this.MINIMUM_CELL_WIDTH = 12;
    this.MINIMUM_CELL_HEIGHT = 12;
    this.currentFocusedCol = 0;
    this.currentFocusedRow = 0;
    this.tbody = null;
    this.rows = null;

    this.SELECTION_MODES =
    {
      RECTANGLE : 0,
      RANDOM    : 1
    };

    this.selectionMode = 0;

    // properties for cell selection
    this.isSelectingCells = false;
    this.selectionStartCell = null;
    this.lastSelectedRegion = null;

    // properties for column selection
    this.isSelectingColumns = false;
    this.selectionStartColumn = null;
    this.lastSelectedColumns = [];

    // properties for row selection
    this.isSelectingRows = false;
    this.selectionStartRow = null;
    this.lastSelectedRows = [];

    // properties needed for horizontal resizing of columns
    this.isResizingHorizontal = false;
    this.resizeOrigXPos = 0;
    this.resizeOrigTH = null;

    // properties needed for vertical resizing of rows
    this.isResizingVertical = false;
    this.resizeOrigYPos = 0;
    this.resizeOrigTD = null;

    // properties for editing
    this.isEditing = false;
    this.spreadsheetWidth = 0;
    this.spreadsheetHeight = 0;


    /**
    * Regular expressions for detecting cell intervals, cells or functions
    */
    this.reInterval = /[a-zA-Z]{1}[0-9]+:[a-zA-Z]{1}[0-9]+/g;
    this.reCell = /[^a-zA-Z]{1}[a-zA-Z]{1}[0-9]+/g;
    this.reFunction = /[a-zA-Z]{3,}[a-zA-Z0-9]*\(/g;

    this.domNode = null;


    /**
     *
     */
    this.months = [ "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec" ];

    this.reDate1 = /[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{1,4}/;
    this.reDate2 = /[0-9]{1,2}-[a-z]{3}-[0-9]{4}/;
    this.reNumber1 = /[^0-9e\.,-]/;
  },




  /*
  *****************************************************************************
     MEMBERS
  *****************************************************************************
  */

  members :
  {
    /**
     * Initializes the graphic widget (table, column/row headers, etc)
     *
     * @type member
     * @param contentPane {var} the ContentPane object corresponding to a tab in a TabContainer
     * @param contentPaneNode {var} TODOC
     * @param formulaField {var} TODOC
     * @param funcButton {var} TODOC
     * @param spread {var} TODOC
     * @param openExisting {var} false if new sheet, true if open existing sheet from xml data
     * @param additionalData {var} if open existing sheet, contains the additional data to be evaluate
     * @return {void} 
     */
    initSheet : function(contentPane, contentPaneNode, formulaField, funcButton, spread, openExisting, additionalData)
    {

      this.spread = spread;

      this.funcButton = funcButton;
      this.formulaField = formulaField;

      this.formulaField.setReadOnly(true);
      this.funcButton.setEnabled(false);

      this.contentPane = contentPane;
      this.domNode = contentPaneNode;

      // Creates new Spread
      if (openExisting == false) {
         // create the table for this sheet
         this.createSpreadsheetCells(contentPane);

         this.resetSpreadsheet();
      }
      // Open an existing Spread
      else {
         this.openSpreadsheet(additionalData);

      }

      // disable selection
      ss.SpreadUtilsHtmlSel.disableSelection(this.domNode);

      this.gainFocus();
      this._focus(this.currentFocusedCol, this.currentFocusedRow);
      this.contentPane.focus();
    },


    /**
     * Called when a sheet is saved and serialized to XML. This retrieves additional DOM cells data that can't be
     * serialized with de "serialize" method. Some code is added to XML. The code can be executed directly by an "Eval" instruction
     * This allows to save all the metadata DOM attributes set is the sheet (.formula, .dwidth....)
     *
     * @type member
     * @return serialAdditData serialized additionnal data (code that can be directly executed by "Eval" instruction)
     */
    getSerialAdditData : function()
    {

      _tdElem = null;
      strBuilder = new qx.util.StringBuilder();
      evalExpres = "";
      for (var i=0; i<this.spreadsheetHeight; i++) {
        

        strBuilder.add("this.rows[" + i + "].dHeight = " + this.rows[i].dHeight + "; ");

        for (var j=0; j<this.spreadsheetWidth; j++) {

//          strBuilder.add("this.rows[" + i + "].cells[" + j + "].firstChild.dWidth = \"" + this.rows[i].cells[j].firstChild.dWidth + "\"; ");
//          strBuilder.add("this.rows[" + i + "].cells[" + j + "].firstChild.style.width = \"" + this.rows[i].cells[j].firstChild.style.width + "\"; ");
//          strBuilder.add("this.rows[" + i + "].cells[" + j + "].firstChild.dHeight = \"" + this.rows[i].cells[j].firstChild.dHeight + "\"; ");
//          strBuilder.add("this.rows[" + i + "].cells[" + j + "].firstChild.style.height = \"" + this.rows[i].cells[j].firstChild.style.height + "\"; ");


          _tdElem = this.getCell(j, i);

          /* Formula */
          if (_tdElem.formula != null) {

//            strBuilder.add("var _tdElem = this.getCell(" + j + ", " + i + "); ");
            strBuilder.add("_tdElem = this.getCell(" + j + ", " + i + "); ");
            strBuilder.add("_tdElem.cellType = \"" + _tdElem.cellType + "\"; ");
            strBuilder.add("_tdElem.formula = \"" + _tdElem.formula + "\"; ");

            /* dWidth */
            if (_tdElem.firstChild.dWidth != null) {
              strBuilder.add("_tdElem.firstChild.dWidth = " + _tdElem.firstChild.dWidth + "; ");
            }
            /* dHeight */
            if (_tdElem.firstChild.dHeight != null) {
              strBuilder.add("_tdElem.firstChild.dHeight = " + _tdElem.firstChild.dHeight + "; ");
            }
            /* dBackgroundColor */
            if (_tdElem.firstChild.dBackgroundColor != null && _tdElem.firstChild.dBackgroundColor != "") {
              strBuilder.add("_tdElem.firstChild.dBackgroundColor = \"" + _tdElem.firstChild.dBackgroundColor + "\"; ");
            }
            /* styleList */
            if (_tdElem.styleList != null && typeof (_tdElem.styleList) != "undefined") {
              strBuilder.add("_tdElem.styleList = new Array(); ");
              strBuilder.add("_tdElem.styleList[this.FORMATTING_TYPES.FONT] = \"" + _tdElem.styleList[this.FORMATTING_TYPES.FONT] + "\"; ");
              strBuilder.add("_tdElem.styleList[this.FORMATTING_TYPES.FONT_SIZE] = \"" + _tdElem.styleList[this.FORMATTING_TYPES.FONT_SIZE] + "\"; ");
              strBuilder.add("_tdElem.styleList[this.FORMATTING_TYPES.COLOR] = \"" + _tdElem.styleList[this.FORMATTING_TYPES.COLOR] + "\"; ");
              strBuilder.add("_tdElem.styleList[this.FORMATTING_TYPES.BG_COLOR] = \"" + _tdElem.styleList[this.FORMATTING_TYPES.BG_COLOR] + "\"; ");
              strBuilder.add("_tdElem.styleList[this.FORMATTING_TYPES.BOLD] = " + _tdElem.styleList[this.FORMATTING_TYPES.BOLD] + "; ");
              strBuilder.add("_tdElem.styleList[this.FORMATTING_TYPES.ITALIC] = " + _tdElem.styleList[this.FORMATTING_TYPES.ITALIC] + "; ");
              strBuilder.add("_tdElem.styleList[this.FORMATTING_TYPES.UNDERLINE] = " + _tdElem.styleList[this.FORMATTING_TYPES.UNDERLINE] + "; ");
              strBuilder.add("_tdElem.styleList[this.FORMATTING_TYPES.ALIGN] = \"" + _tdElem.styleList[this.FORMATTING_TYPES.ALIGN] + "\"; ");
            }
          }
        }


      }


      /* Column headers */
      
      ths = this.domNode.getElementsByTagName("TH");
      for (var i=1; i<ths.length; i++)
      {
              strBuilder.add("ths[" + i + "].dCellIndex = " + i + "- 1; ");
              strBuilder.add("ths[" + i + "].dWidth =  " + ths[i].dWidth + "; ");
      }


      evalExpres = strBuilder.get();
      return evalExpres;

    },



    /**
     * Called when a sheet is losing focus (other sheet becomes active) - disconnects event handlers
     *
     * @type member
     * @return {void} 
     */
    loseFocus : function()
    {
      this.contentPane.removeEventListener("mousedown", this.onMouseDown, this);
      this.contentPane.removeEventListener("mousemove", this.onMouseMove, this);
      this.contentPane.removeEventListener("mouseup", this.onMouseUp, this);
      this.contentPane.removeEventListener("keypress", this.onKeyPress, this);
      this.formulaField.removeEventListener("keypress", this.onKeyPress, this);
      this.contentPane.removeEventListener("dblclick", this.onDblClick, this);
    },


    /**
     * Called when a sheet is gaining focus - reconnects event handlers
     *
     * @type member
     * @return {void} 
     */
    gainFocus : function()
    {
      // make sure events are not connected twice
      this.loseFocus();
      this.contentPane.addEventListener("mousedown", this.onMouseDown, this);
      this.contentPane.addEventListener("mousemove", this.onMouseMove, this);
      this.contentPane.addEventListener("mouseup", this.onMouseUp, this);
      this.contentPane.addEventListener("keypress", this.onKeyPress, this);
      this.formulaField.addEventListener("keypress", this.onKeyPress, this);
      this.contentPane.addEventListener("dblclick", this.onDblClick, this);
    },


    /**
     * Creates the spreadsheet cells (the actual <TABLE> element)
     *
     * @type member
     * @param contentPane {var} the ContentPane object associated with a TabContainer
     * @return {void} 
     */
    createSpreadsheetCells : function(contentPane)
    {
      
      var strBuilder = new qx.util.StringBuilder();
      var html = "";

      strBuilder.add("<table id=\"" + this.sheetName + "\" class=\"sheet\" cellspacing=\"0\" cellpadding=\"0\" >");
      strBuilder.add("  <thead>");
      strBuilder.add("    <th class=\"sheetRow1stCell\"></th>");

      for (var col=0; col<this.initColsCount; col++) {
        strBuilder.add("    <th><div class=\"header\"><div class=\"headerText\">A</div><div class=\"horizontalResizer\">&nbsp;</div></div></th>");
      }

      strBuilder.add("  </thead>");
      strBuilder.add("  <tbody>");

      for (var row=0; row<this.initRowsCount; row++)
      {
        strBuilder.add("    <tr>");
        strBuilder.add("      <td class=\"sheetRow1stCell\"><div><div class=\"rowHeaderText\"></div><div class=\"verticalResizer\">&nbsp;</div></div></td>");

        for (var col=0; col<this.initColsCount; col++) {
          strBuilder.add("      <td class=\"sheetCell\"><div class=\"sheetCellContent\"></div></td>");
        }

        strBuilder.add("    </tr>");
      }

      strBuilder.add("  </tbody>");
      strBuilder.add("</table>");

      html = strBuilder.get();

      contentPane.setHtml(html);
    },


    /**
     * Applies labels to column/row headers. Fixes width of first column and initializes internal data for each cell/header
     *
     * @type member
     * @return {void} 
     */
    resetSpreadsheet : function()
    {
      this.tbody = this.domNode.getElementsByTagName("tbody")[0];
      this.rows = this.tbody.getElementsByTagName("tr");

      for (var i=0; i<this.rows.length; i++)
      {

        this.rows[i].cells[0].firstChild.firstChild.innerHTML = "" + (i + 1);


        this.rows[i].dHeight = 20;  // ss.SpreadUtilsHtmlLay.getBorderBoxHeight(this.rows[i]).height;


        for (var j=0; j<this.rows[i].cells.length; j++)
        {
          // OSY
          var tmpW = 65;
          var tmpH = 16;

          this.rows[i].cells[j].firstChild.dWidth = tmpW;
          this.rows[i].cells[j].firstChild.style.width = tmpW;
          this.rows[i].cells[j].firstChild.dHeight = tmpH;
          this.rows[i].cells[j].firstChild.style.height = tmpH;
        }


        var eldiv = this.rows[i].cells[0].getElementsByTagName("div");

        for (var ss=0; ss<eldiv.length; ss++) {
          eldiv[ss].style.width = "17px";
        }

        this.rows[i].cells[0].style.width = "20px";
        this.rows[i].cells[0].firstChild.style.height = "20px";
        this.rows[i].cells[0].firstChild.firstChild.style.height = "18px";
        this.rows[i].cells[0].firstChild.firstChild.nextSibling.style.height = "2px";
      }


      var ths = this.domNode.getElementsByTagName("TH");
      ths[0].style.width = "65px";

      for (var i=1; i<ths.length; i++)
      {
        ths[i].dCellIndex = i - 1;

        // OSY
        ths[i].dWidth = 65;  // ss.SpreadUtilsHtmlLay.getBorderBoxWidth(ths[i]).width;

        if (ths[i].firstChild && ths[i].firstChild.firstChild) {
          ths[i].firstChild.firstChild.innerHTML = String.fromCharCode("A".charCodeAt(0) + i - 1);
        }
      }


      var client = qx.core.Client.getInstance();

      this.spreadsheetWidth = this.rows[0].cells.length - 1;
      this.spreadsheetHeight = client.isMshtml() ? (this.rows.length - 1) : this.rows.length;
    },


    /**
     * Called when a sheet is opened. Allows to initialize the DOM objects of the opened sheet.
     * Serialized data is evaluated by the Eval instruction.
     *
     * @param additionalData {var} contains the additional data to be evaluate for opening existing sheet
     * @type member
     *
     */
    openSpreadsheet : function(additionalData)
    {
         var client = qx.core.Client.getInstance();
         this.tbody = this.domNode.getElementsByTagName("tbody")[0];
         this.rows = this.tbody.getElementsByTagName("tr");
         
         ths = this.domNode.getElementsByTagName("TH");

         this.debug("Additional Data will be executed : " + additionalData);
         result = eval(additionalData);




         this.spreadsheetWidth = this.rows[0].cells.length - 1;
         this.spreadsheetHeight = client.isMshtml() ? (this.rows.length - 1) : this.rows.length;
    },


    /**
     * Refresh spreadsheet after insert, del... rows/columns
     *
     * @type member
     * @return {void} 
     */
    refreshSpreadsheet : function()
    {
      this.tbody = this.domNode.getElementsByTagName("tbody")[0];
      this.rows = this.tbody.getElementsByTagName("tr");

      for (var i=0; i<this.rows.length; i++)
      {
        this.rows[i].cells[0].firstChild.firstChild.innerHTML = "" + (i + 1);
      }


      ths = this.domNode.getElementsByTagName("TH");

      for (var i=1; i<ths.length; i++)
      {
        ths[i].dCellIndex = i - 1;


        if (ths[i].firstChild && ths[i].firstChild.firstChild) {
          ths[i].firstChild.firstChild.innerHTML = String.fromCharCode("A".charCodeAt(0) + i - 1);
        }
      }


      client = qx.core.Client.getInstance();

      this.spreadsheetWidth = this.rows[0].cells.length - 1;
      this.spreadsheetHeight = client.isMshtml() ? (this.rows.length - 1) : this.rows.length;
    },




    /**
     * Handles onmousedown which can cause several relevant events for spreadsheet: focusing, selection, resizing
     *
     * @type member
     * @param e {Event} TODOC
     * @return {void} 
     */
    onMouseDown : function(e)
    {
      this.deselectAll();

      var currentTDElem = ss.SpreadUtilsHtmlCommon.getParentByType(e.getDomTarget(), "TD");

      if (currentTDElem && ss.SpreadUtilsHtmlStyle.hasClass(currentTDElem, "sheetCell"))
      {
        this.isSelectingCells = true;
        this.selectionStartCell = currentTDElem;
        this.selectCellByTD(this.selectionStartCell, true);
        e.setDefaultPrevented(true);
      }

      var currentDIVElem = ss.SpreadUtilsHtmlCommon.getParentByType(e.getDomTarget(), "DIV");

      if (currentDIVElem && ss.SpreadUtilsHtmlStyle.hasClass(currentDIVElem, "horizontalResizer"))
      {
        this.isResizingHorizontal = true;
        var pos = ss.SpreadUtilsHtmlCommon.getCursorPosition(e);
        this.resizeOrigXPos = pos.x;
        this.resizeOrigTH = ss.SpreadUtilsHtmlCommon.getParentByType(e.getDomTarget(), "TH");
      }

      if (currentDIVElem && ss.SpreadUtilsHtmlStyle.hasClass(currentDIVElem, "verticalResizer"))
      {
        this.isResizingVertical = true;
        var pos = ss.SpreadUtilsHtmlCommon.getCursorPosition(e);
        this.resizeOrigYPos = pos.y;
        this.resizeOrigTD = ss.SpreadUtilsHtmlCommon.getParentByType(e.getDomTarget(), "TD");
      }

      if (!this.isResizingHorizontal && !this.isResizingVertical)
      {
        var currentTH = ss.SpreadUtilsHtmlCommon.getParentByType(e.getDomTarget(), "TH");

        if (currentTH)
        {
          this.isSelectingColumns = true;
          this.selectionStartColumn = currentTH.dCellIndex;
          this.selectColumns(this.selectionStartColumn);
        }

        var currentTDRowHeader = ss.SpreadUtilsHtmlCommon.getParentByType(e.getDomTarget(), "TD");

        if (currentTDRowHeader && ss.SpreadUtilsHtmlStyle.hasClass(currentTDRowHeader, "sheetRow1stCell"))
        {
          this.isSelectingRows = true;
          this.selectionStartRow = this.getCellRow(currentTDRowHeader);
          this.selectRows(this.selectionStartRow);
        }
      }
    },


    /**
     * Handles onmousemove - useful when resizing to dynamically resize as mouse moves
     *
     * @type member
     * @param e {Event} TODOC
     * @return {void} 
     */
    onMouseMove : function(e)
    {
      if (this.isResizingHorizontal) {
        this.resizeCol(e, false);
      } else if (this.isResizingVertical) {
        this.resizeRow(e, false);
      }
      else if (this.isSelectingCells)
      {
        var currentTDElem = ss.SpreadUtilsHtmlCommon.getParentByType(e.getDomTarget(), "TD");

        if (currentTDElem && ss.SpreadUtilsHtmlStyle.hasClass(currentTDElem, "sheetCell"))
        {
          this.deselectAll();

          this.selectRegion(this.getCellCol(currentTDElem), this.getCellCol(this.selectionStartCell), this.getCellRow(currentTDElem), this.getCellRow(this.selectionStartCell), true);
        }
      }
      else if (this.isSelectingColumns)
      {
        // the mouse may move over THs or TDs. However we should detect in each case the column
        var currentTDElem = ss.SpreadUtilsHtmlCommon.getParentByType(e.getDomTarget(), "TD");

        if (currentTDElem && ss.SpreadUtilsHtmlStyle.hasClass(currentTDElem, "sheetCell"))
        {
          var selectionEndColumn = this.getCellCol(currentTDElem);
          this.selectColumns(selectionEndColumn);
        }
        else
        {
          var currentTHElem = ss.SpreadUtilsHtmlCommon.getParentByType(e.getDomTarget(), "TH");

          if (currentTHElem)
          {
            var selectionEndColumn = currentTHElem.dCellIndex;
            this.selectColumns(selectionEndColumn);
          }
        }
      }
      else if (this.isSelectingRows)
      {
        var currentTDElem = ss.SpreadUtilsHtmlCommon.getParentByType(e.getDomTarget(), "TD");

        if (currentTDElem && (ss.SpreadUtilsHtmlStyle.hasClass(currentTDElem, "sheetCell") || ss.SpreadUtilsHtmlStyle.hasClass(currentTDElem, "sheetRow1stCell")))
        {
          var selectionEndRow = this.getCellRow(currentTDElem);
          this.selectRows(selectionEndRow);
        }
      }
    },


    /**
     * Handles onmouseup - either selection of resizing have finished
     *
     * @type member
     * @param e {Event} TODOC
     * @return {void} 
     */
    onMouseUp : function(e)
    {
      if (this.isSelectingCells)
      {
        // unselect only if the mouse up occured on the same cell
        if (this.selectionStartCell && this.selectionStartCell == ss.SpreadUtilsHtmlCommon.getParentByType(e.getDomTarget(), "TD"))
        {
          this.selectCellByTD(this.selectionStartCell, false);
          this.focusOn(this.selectionStartCell);
        }

        this.isSelectingCells = false;
      }
      else if (this.isResizingHorizontal)
      {
        this.resizeCol(e, true);
        this.isResizingHorizontal = false;
      }
      else if (this.isResizingVertical)
      {
        this.resizeRow(e, true);
        this.isResizingVertical = false;
      }
      else if (this.isSelectingColumns)
      {
        this.isSelectingColumns = false;
      }
      else if (this.isSelectingRows)
      {
        this.isSelectingRows = false;
      }
    },


    /**
     * Ondblclick - the editing box is shown
     *
     * @type member
     * @param e {Event} TODOC
     * @return {void} 
     */
    onDblClick : function(e)
    {
      var currentTDElem = ss.SpreadUtilsHtmlCommon.getParentByType(e.getDomTarget(), "TD");

      if (currentTDElem && ss.SpreadUtilsHtmlStyle.hasClass(currentTDElem, "sheetCell"))
      {
        this.deselectAll();
        this.editFormulaField(currentTDElem, true);
      }
    },


    /**
     * Onkeypress - manages delete, enter, backspace, esc, space keys
     *
     * @type member
     * @param e {Event} TODOC
     * @return {void} 
     */
    onKeyPress : function(e)
    {
      var keyCode = e.getKeyIdentifier();

      var keyHandled = false;

      if (keyCode == "Down")
      {
        if (!this.isEditing)
        {
          this.moveFocus(0, 1);
          keyHandled = true;
        }
      }
      else if (keyCode == "Up")
      {
        if (!this.isEditing)
        {
          this.moveFocus(0, -1);
          keyHandled = true;
        }
      }
      else if (keyCode == "Left")
      {
        if (!this.isEditing)
        {
          this.moveFocus(-1, 0);
          keyHandled = true;
        }
      }
      else if (keyCode == "Right" || keyCode == "Tab")
      {
        if (!this.isEditing)
        {
          this.moveFocus(1, 0);
          keyHandled = true;
        }
      }
      else if (keyCode == "Escape")
      {
        if (this.isEditing)
        {
          this.exitFormulaEdit(true);
          keyHandled = true;
        }
      }
      else if (keyCode == "Backspace" || keyCode == "Delete")
      {
        if (!this.isEditing) {
          this.eraseCurrentCellContent();
        }
      }
      else if (keyCode == "Enter" || keyCode == "F2")
      {
        if (!this.isEditing)
        {
          var tdElem = this.getCell(this.currentFocusedCol, this.currentFocusedRow);
          this.editFormulaField(tdElem, true);
          keyHandled = true;
        }
        else
        {
          this.moveFocus(0, 1);
          keyHandled = true;
        }
      }

      /*
       * others cases ==> Editing mode start and the formula field 
       * is initialized with current keyboard input value 
       */

      else if (!this.isEditing && e.getCharCode() != 0 && !e.isCtrlOrCommandPressed() && !e.isAltPressed())
      {
        this.formulaField.focus();
        var tdElem = this.getCell(this.currentFocusedCol, this.currentFocusedRow);
        tdElem.formula = keyCode;
        this.editFormulaField(tdElem, false);
      }

      if (keyHandled == true) {}
    },

    //    	e.preventDefault();
    //    	e.stopPropagation();
    /**
     * Moves focus to the specified cell
     *
     * @type member
     * @param deltaX {var} TODOC
     * @param deltaY {var} TODOC
     * @return {void} 
     */
    moveFocus : function(deltaX, deltaY)
    {
      var newCol = this.currentFocusedCol;
      var anyChange = false;

      if (newCol + deltaX < this.spreadsheetWidth && newCol + deltaX >= 0)
      {
        newCol = newCol + deltaX;
        anyChange = true;
      }

      var newRow = this.currentFocusedRow;

      if (newRow + deltaY < this.spreadsheetHeight && newRow + deltaY >= 0)
      {
        newRow = newRow + deltaY;
        anyChange = true;
      }

      if (anyChange)
      {
        this.unfocus(this.currentFocusedCol, this.currentFocusedRow);
        this.currentFocusedCol = newCol;
        this.currentFocusedRow = newRow;
        this._focus(this.currentFocusedCol, this.currentFocusedRow);
      }
    },


    /**
     * Focuses on a given element
     *
     * @type member
     * @param elem {Element} the <TD> element to focus on
     * @return {void} 
     */
    focusOn : function(elem)
    {
      var _col = this.getCellCol(elem);
      var _row = this.getCellRow(elem);

      if (_row < this.spreadsheetHeight && _col < this.spreadsheetWidth) {
        this.focusAt(_col, _row);
      }
    },


    /**
     * Focuses at a given location
     *
     * @type member
     * @param _col {var} the col coordinate where to focus
     * @param _row {var} the row coordinate where to focus
     * @return {void} 
     */
    focusAt : function(_col, _row)
    {
      if (_row >= this.spreadsheetHeight) {
        _row = this.spreadsheetHeight - 1;
      }

      if (_col >= this.spreadsheetWidth) {
        _col = this.spreadsheetWidth - 1;
      }

      this.unfocus(this.currentFocusedCol, this.currentFocusedRow);
      this.currentFocusedCol = _col;
      this.currentFocusedRow = _row;
      this.selectionStartCell = this.getCell(_col, _row);
      this._focus(this.currentFocusedCol, this.currentFocusedRow);
    },


    /**
     * Internal fn which actually does the focusing
     *
     * @type member
     * @param _col {var} the col coordinate where to focus
     * @param _row {var} the row coordinate where to focus
     * @return {void} 
     */
    _focus : function(_col, _row)
    {
      if (typeof (_col) != "undefined" && typeof (_row) != "undefined")
      {
        if (_row < this.spreadsheetHeight && _col < this.spreadsheetWidth)
        {
          var _tdElem = this.getCell(_col, _row);

          var _italic = false;
          var _bold = false;
          var _underline = false;
          var _align = "";

          if (typeof (_tdElem.styleList) != "undefined")
          {
            _bold = _tdElem.styleList[this.FORMATTING_TYPES.BOLD] ? _tdElem.styleList[this.FORMATTING_TYPES.BOLD] : false;
            _italic = _tdElem.styleList[this.FORMATTING_TYPES.ITALIC] ? _tdElem.styleList[this.FORMATTING_TYPES.ITALIC] : false;
            _underline = _tdElem.styleList[this.FORMATTING_TYPES.UNDERLINE] ? _tdElem.styleList[this.FORMATTING_TYPES.UNDERLINE] : false;
            _align = _tdElem.styleList[this.FORMATTING_TYPES.ALIGN];
          }

          // this._fireEvent("setFormatting", _bold, _italic, _underline);
          // this.createDispatchDataEvent("setFormatting", this.spread, this, _bold, _italic, _underline);
          this.spread.styleBold.setChecked(_bold);
          this.spread.styleItalic.setChecked(_italic);
          this.spread.styleUnderline.setChecked(_underline);

          this.spread.alignLeft.setChecked(false);
          this.spread.alignCenter.setChecked(false);
          this.spread.alignRight.setChecked(false);
          this.spread.alignJustify.setChecked(false);

          if (_align == "left") {
            this.spread.alignLeft.setChecked(true);
          } else if (_align == "center") {
            this.spread.alignCenter.setChecked(true);
          } else if (_align == "right") {
            this.spread.alignRight.setChecked(true);
          } else if (_align == "justify") {
            this.spread.alignJustify.setChecked(true);
          }

          var w = _tdElem.firstChild.dWidth;
          var h = _tdElem.firstChild.dHeight;

          // OSY
          // _tdElem.firstChild.style.width = this.toPx(w - (qx.core.Client.getInstance().isMshtml() ? 0 : 4));
          // _tdElem.firstChild.style.height = this.toPx(h - (qx.core.Client.getInstance().isMshtml() ? 0 : 4));
          // _tdElem.firstChild.style.width = this.toPx(w);
          // _tdElem.firstChild.style.height = this.toPx(h);
          ss.SpreadUtilsHtmlStyle.addClass(_tdElem.firstChild, "focused");
        }
      }
    },


    /**
     * Unfocuses the cell at a given position
     *
     * @type member
     * @param _col {var} the col coordinate where to unfocus
     * @param _row {var} the row coordinate where to unfocus
     * @return {void} 
     */
    unfocus : function(_col, _row)
    {
      if (this.isEditing) {
        this.exitFormulaEdit();
      }

      if (typeof (_col) != "undefined" && typeof (_row) != "undefined")
      {
        if (_row < this.spreadsheetHeight && _col < this.spreadsheetWidth)
        {
          var _tdElem = this.getCell(_col, _row);

          var w = _tdElem.firstChild.dWidth;
          var h = _tdElem.firstChild.dHeight;

          // OSY
          // _tdElem.firstChild.style.width = this.toPx(w);
          // _tdElem.firstChild.style.height = this.toPx(h);
          ss.SpreadUtilsHtmlStyle.removeClass(_tdElem.firstChild, "focused");
        }
      }
    },


    /**
     * Selects a region of the spreadsheet
     *
     * @type member
     * @param col1 {var} the start col
     * @param col2 {var} the end col
     * @param row1 {var} the start row
     * @param row2 {var} the end row
     * @param _select {var} whether to select or deselect the region
     * @return {void} 
     */
    selectRegion : function(col1, col2, row1, row2, _select)
    {
      var minCol = col1 < col2 ? col1 : col2;
      var maxCol = col1 > col2 ? col1 : col2;
      var minRow = row1 < row2 ? row1 : row2;
      var maxRow = row1 > row2 ? row1 : row2;

      this.lastSelectedRegion = [ minCol, maxCol, minRow, maxRow ];

      for (var i=minCol; i<=maxCol; i++)
      {
        for (var j=minRow; j<=maxRow; j++) {
          this.selectCell(i, j, _select);
        }
      }

      if (this.selectionStartCell) {
        this.focusOn(this.selectionStartCell);
      }

      if (!_select) {
        this.lastSelectedRegion = null;
      }
    },


    /**
     * Selects a cell given the HTML element
     *
     * @type member
     * @param _tdElem {var} the <TD> element to select
     * @param _select {var} whether to select or deselect the HTML element
     * @return {void} 
     */
    selectCellByTD : function(_tdElem, _select)
    {
      if (typeof (_tdElem) != "undefined") {
        this.selectCell(this.getCellCol(_tdElem), this.getCellRow(_tdElem), _select);
      }
    },


    /**
     * Selects or deselects a cell. Selection means changing the color to a lighter/darker blue depending on B's value (B from RGB)
     *
     * @type member
     * @param _col {var} the column (0-n, where 0 corresponds to the A column of the spreadsheet which is actually the second column of the table
     * @param _row {var} the row (0-n, 0 - corresponds similarly to row 1)
     * @param _select {var} whether to select or deselect the cell
     * @return {void} 
     */
    selectCell : function(_col, _row, _select)
    {
      if (typeof (_col) != "undefined" && typeof (_row) != "undefined")
      {
        if (_row < this.spreadsheetHeight && _col < this.spreadsheetWidth)
        {
          var tdElem = this.getCell(_col, _row);

          if (_select)
          {
            var bgColor = ss.SpreadUtilsHtmlStyle.getBackgroundColor(tdElem.firstChild);

            bgColor[0] = parseInt(bgColor[0]);
            bgColor[1] = parseInt(bgColor[1]);
            bgColor[2] = parseInt(bgColor[2]);
            bgColor[0] = bgColor[0] > 128 ? bgColor[0] - 32 : bgColor[0] + 32;
            bgColor[1] = bgColor[1] > 128 ? bgColor[1] - 32 : bgColor[1] + 32;
            tdElem.firstChild.style.backgroundColor = ss.SpreadUtilsHtmlStyle.rgb2hex(bgColor[0], bgColor[1], bgColor[2]);
          }
          else
          {
            // OSY ==> deactiv. for cell color selection
            tdElem.firstChild.dBackgroundColor = "";
            tdElem.firstChild.style.backgroundColor = tdElem.firstChild.dBackgroundColor;
          }
        }
      }
    },


    /**
     * Selects an entire column
     *
     * @type member
     * @param idx {var} the index of the column to be selected
     * @param _select {var} whether to select or deselect the column
     * @return {void} 
     */
    selectColumn : function(idx, _select)
    {
      for (var i=0; i<this.spreadsheetHeight; i++) {
        this.selectCell(idx, i, _select);
      }
    },


    /**
     * Selects an interval of columns
     *
     * @type member
     * @param selectionEndColumn {var} TODOC
     * @return {void} 
     */
    selectColumns : function(selectionEndColumn)
    {
      this.unfocus(this.currentFocusedCol, this.currentFocusedRow);
      this.deselectAll();
      this.lastSelectedColumns = [];
      var minCol = selectionEndColumn < this.selectionStartColumn ? selectionEndColumn : this.selectionStartColumn;
      var maxCol = selectionEndColumn > this.selectionStartColumn ? selectionEndColumn : this.selectionStartColumn;

      for (var i=minCol; i<=maxCol; i++)
      {
        this.addColumnToSelection(i);
        this.selectColumn(i, true);
      }

      // select first cell of first columnx
      this.focusAt(this.selectionStartColumn, 0);
    },


    /**
     * Adds a column to the list of columns to be selected
     *
     * @type member
     * @param idx {var} the index of the column to be selected
     * @return {void} 
     */
    addColumnToSelection : function(idx) {
      this.lastSelectedColumns[this.lastSelectedColumns.length] = idx;
    },


    /**
     * Selects an entire row
     *
     * @type member
     * @param idx {var} the index of the row to be selected
     * @param _select {var} whether to select or deselect the row
     * @return {void} 
     */
    selectRow : function(idx, _select)
    {
      for (var i=0; i<this.spreadsheetWidth; i++) {
        this.selectCell(i, idx, _select);
      }
    },


    /**
     * Selects an interval of rows
     *
     * @type member
     * @param selectionEndRow {var} TODOC
     * @return {void} 
     */
    selectRows : function(selectionEndRow)
    {
      this.deselectAll();
      this.lastSelectedRows = [];
      var minRow = selectionEndRow < this.selectionStartRow ? selectionEndRow : this.selectionStartRow;
      var maxRow = selectionEndRow > this.selectionStartRow ? selectionEndRow : this.selectionStartRow;

      for (var i=minRow; i<=maxRow; i++)
      {
        this.addRowToSelection(i);
        this.selectRow(i, true);
      }

      // select first cell of first row
      this.focusAt(1, this.selectionStartRow);
    },


    /**
     * Adds a row to the list of rows to be selected
     *
     * @type member
     * @param idx {var} the index of the row to be selected
     * @return {void} 
     */
    addRowToSelection : function(idx) {
      this.lastSelectedRows[this.lastSelectedRows.length] = idx;
    },


    /**
     * Returns the selection mode. Possible values are: SELECTION_MODES.RECTANGLE and SELECTION_MODES.RANDOM
     * Random mode will be used when selecting random cells using SHIFT & CTRL. Not supported yet.
     *
     * @type member
     * @return {var} the selection mode
     */
    getSelectionMode : function() {
      return this.selectionMode;
    },


    /**
     * Returns a list of selected cells - to be used when selection mode is SELECTION_MODES.RANDOM
     *
     * @type member
     * @return {var} a list of selected cells
     */
    getSelection : function()
    {
      var selectedCells = new Array();

      if (this.getSelectionMode() == this.SELECTION_MODES.RECTANGLE)
      {
        if (this.lastSelectedRegion != null)
        {
          for (var i=this.lastSelectedRegion[0]; i<=this.lastSelectedRegion[1]; i++)
          {
            for (var j=this.lastSelectedRegion[2]; j<=this.lastSelectedRegion[3]; j++) {
              selectedCells[selectedCells.length] = [ i, j ];
            }
          }
        }
        else
        {
          selectedCells[selectedCells.length] = [ this.currentFocusedCol, this.currentFocusedRow ];
        }
      }

      return selectedCells;
    },


    /**
     * Returns the selection rectangle - to be used when selection mode is SELECTION_MODES.RECTANGLE
     *
     * @type member
     * @return {var | Array} an array with 4 coordinates: [colStart, colEnd, rowStart, rowEnd]
     */
    getSelectionRectangle : function()
    {
      if (this.lastSelectedRegion && this.lastSelectedRegion != null) {
        return this.lastSelectedRegion;
      } else {
        return [ this.currentFocusedCol, this.currentFocusedCol, this.currentFocusedRow, this.currentFocusedRow ];
      }
    },


    /**
     * Deselects everything
     *
     * @type member
     * @return {void} 
     */
    deselectAll : function()
    {
      this.unfocus(this.currentFocusedCol, this.currentFocusedRow);

      this.resetSelectedColumns();
      this.resetSelectedRows();

      if (this.lastSelectedRegion && this.lastSelectedRegion != null) {
        this.selectRegion(this.lastSelectedRegion[0], this.lastSelectedRegion[1], this.lastSelectedRegion[2], this.lastSelectedRegion[3], false);
      }
    },


    /**
     * Deselects selected columns if any
     *
     * @type member
     * @return {void} 
     */
    resetSelectedColumns : function()
    {
      for (var i=0; i<this.lastSelectedColumns.length; i++) {
        this.selectColumn(this.lastSelectedColumns[i], false);
      }

      this.lastSelectedColumns = [];
    },


    /**
     * Deselects selected rows if any
     *
     * @type member
     * @return {void} 
     */
    resetSelectedRows : function()
    {
      for (var i=0; i<this.lastSelectedRows.length; i++) {
        this.selectRow(this.lastSelectedRows[i], false);
      }

      this.lastSelectedRows = [];
    },


    /**
     * Resizes a column
     *
     * @type member
     * @param e {Event} the DOM event
     * @param storePosition {var} whether to store position or not (used when the mouse is still moving)
     * @return {void} 
     */
    resizeCol : function(e, storePosition)
    {
      var pos = ss.SpreadUtilsHtmlCommon.getCursorPosition(e);
      var newX = pos.x;

      var currentCol = this.resizeOrigTH.dCellIndex + 1;

      // if(currentResizedCol) {
      var delta = newX - this.resizeOrigXPos;
      var newWidth = this.resizeOrigTH.dWidth + delta;

      if (newWidth < this.MINIMUM_CELL_WIDTH) {
        newWidth = this.MINIMUM_CELL_WIDTH;
      }

      this.resizeOrigTH.style.width = this.toPx(newWidth);
      this.resizeOrigTH.firstChild.style.width = this.toPx(newWidth - 3);
      this.resizeOrigTH.firstChild.firstChild.style.width = this.toPx(newWidth - 5);

      for (var i=0; i<this.rows.length; i++) {
        this.rows[i].cells[currentCol].firstChild.style.width = this.toPx(newWidth);
      }

      if (storePosition)
      {
        this.resizeOrigTH.dWidth = newWidth;

        for (var i=0; i<this.rows.length; i++) {
          this.rows[i].cells[currentCol].firstChild.dWidth = newWidth;
        }

        this._focus(this.currentFocusedCol, this.currentFocusedRow);
      }
    },

    // }
    /**
     * Resizes a row
     *
     * @type member
     * @param e {Event} the DOM event
     * @param storePosition {var} whether to store position or not (used when the mouse is still moving)
     * @return {void} 
     */
    resizeRow : function(e, storePosition)
    {
      var pos = ss.SpreadUtilsHtmlCommon.getCursorPosition(e);
      var newY = pos.y;

      var currentRow = this.getCellRow(this.resizeOrigTD);

      var delta = newY - this.resizeOrigYPos;
      var newHeight = this.rows[currentRow].dHeight + delta;

      if (newHeight < this.MINIMUM_CELL_HEIGHT) {
        newHeight = this.MINIMUM_CELL_HEIGHT;
      }

      // resize 1st cell (number + vertical resizer)
      this.rows[currentRow].cells[0].style.height = this.toPx(newHeight);
      this.rows[currentRow].cells[0].firstChild.style.height = this.toPx(newHeight);
      this.rows[currentRow].cells[0].firstChild.firstChild.style.height = this.toPx(newHeight - 2);  //  - (qx.core.Client.getInstance().isMshtml() ? 2 : 0));

      for (var i=1; i<this.rows[0].cells.length; i++) {
        this.rows[currentRow].cells[i].firstChild.style.height = this.toPx(newHeight);
      }

      if (storePosition)
      {
        this.rows[currentRow].dHeight = newHeight;

        for (var i=1; i<this.rows[0].cells.length; i++) {
          this.rows[currentRow].cells[i].firstChild.dHeight = newHeight;
        }

        this._focus(this.currentFocusedCol, this.currentFocusedRow);
      }
    },


    /**
     * Displays the input element at a given location
     *
     * @type member
     * @param _tdElem {var} the <TD> element over which the input element has to be displayed
     * @param _textSelect {var} TODOC
     * @return {void} 
     */
    editFormulaField : function(_tdElem, _textSelect)
    {
      this.isEditing = true;

      this.funcButton.setEnabled(true);

      with (this.formulaField)
      {
        setReadOnly(false);

        focus();

        tdElem = _tdElem;

        if (tdElem != null)
        {
          if (typeof (tdElem.formula) == "undefined") {
            setValue("");
          }
          else
          {
            var formValue = "" + _tdElem.formula;

            // replace dot by the decimal separator
            if (typeof (_tdElem.formula) == "number") {
              formValue = formValue.replace(/[.]+/, this.DECIMAL_SEP);
            }

            setValue(formValue);
          }
        }

        /* if(tdElem.cellType == this.CELL_TYPES.FORMULA) {
            value = tdElem.formula;
        } else {
            value = tdElem.firstChild.innerHTML;
        } */

        else
        {
          setValue("");
        }

        oldValue = getValue();

        if (typeof (oldValue) == "undefined" || oldValue == "undefined") {
          oldValue = "";
        }

        if (_textSelect)
        {
          setSelectionStart(0);
          setSelectionLength(getValue().length);
        }
        else
        {
          setSelectionStart(getValue().length);
          setSelectionLength(getValue().length);
        }
      }
    },


    /**
     * Called when the formula field input is canceled or validated
     *
     * @type member
     * @param isCancel {var} whether to cancel or not the modifications to the cell
     * @return {void} 
     */
    exitFormulaEdit : function(isCancel)
    {
      // OSY
      this.isEditing = false;
      this.contentPane.focus();

      with (this.formulaField)
      {
        // prevent this from executing twice - change the value and hide the input only if input is visible
        if (!this.formulaField.getReadOnly())
        {
          if (tdElem != null)
          {
            if (isCancel) {
              tdElem.formula = oldValue;
            }
            else
            {
              tdElem.formula = getValue();
              this.evalFormula(tdElem, true);
            }
          }

          setValue("");

          this.formulaField.setReadOnly(true);
          this.funcButton.setEnabled(false);
        }
      }
    },


    /**
     * Deletes the content of the current focused cell
     *
     * @type member
     * @return {void} 
     */
    eraseCurrentCellContent : function()
    {
      var tdElem = this.getCell(this.currentFocusedCol, this.currentFocusedRow);
      tdElem.firstChild.innerHTML = "";
    },


    /**
     * Inserts a new column before the column of the current focused cell
     *
     * @type member
     * @return {void} 
     */
    insertColumnBefore : function()
    {
      try
      {
        var columnToInsertBefore = this.currentFocusedCol + 1;
        var thead = this.domNode.getElementsByTagName("THEAD")[0];
        var ths = thead.getElementsByTagName("TH");
        var newTH = ths[columnToInsertBefore].cloneNode(true);

        ss.SpreadUtilsHtmlDom.insertBefore(newTH, ths[columnToInsertBefore]);

        for (var i=0; i<this.rows.length; i++)
        {
          var newTD = this.rows[i].cells[columnToInsertBefore].cloneNode(true);
          ss.SpreadUtilsHtmlDom.insertBefore(newTD, this.rows[i].cells[columnToInsertBefore]);
        }

        this.refreshSpreadsheet();

        // reset cell's formatting and content
        if (columnToInsertBefore >= 1)
        {
          for (var i=0; i<this.rows.length; i++)
          {
            this.rows[i].cells[columnToInsertBefore].firstChild.innerHTML = "";
            this.unfocus(columnToInsertBefore, i);
            this.selectCell(columnToInsertBefore, i, false);
          }
        }
      }
      catch(e)
      {
        alert(e + " - " + e.name + " - " + e.message);
      }
    },


    /**
     * Inserts a new column after the column of the current focused cell
     *
     * @type member
     * @return {void} 
     */
    insertColumnAfter : function()
    {
      try
      {
        var thead = this.domNode.getElementsByTagName("THEAD")[0];
        var ths = thead.getElementsByTagName("TH");
        var newTH = ths[this.currentFocusedCol].cloneNode(true);

        ss.SpreadUtilsHtmlDom.insertAfter(newTH, ths[this.currentFocusedCol]);

        for (var i=0; i<this.rows.length; i++)
        {
          var newTD = this.rows[i].cells[this.currentFocusedCol].cloneNode(true);
          ss.SpreadUtilsHtmlDom.insertAfter(newTD, this.rows[i].cells[this.currentFocusedCol]);
        }

        this.refreshSpreadsheet();

        // reset cell's formatting and content
        if (this.currentFocusedCol >= 1)
        {
          for (var i=0; i<this.rows.length; i++)
          {
            this.rows[i].cells[this.currentFocusedCol + 1].firstChild.innerHTML = "";
            this.unfocus(this.currentFocusedCol + 1, i);
            this.selectCell(this.currentFocusedCol + 1, i, false);
          }
        }
      }
      catch(e)
      {
        alert(e + " $$ " + e.name + " $$ " + e.message);
      }
    },


    /**
     * Inserts a new row before the row of the current focused cell
     *
     * @type member
     * @return {void} 
     */
    insertRowBefore : function()
    {
      try
      {
        var newTR = this.rows[this.currentFocusedRow].cloneNode(true);
        ss.SpreadUtilsHtmlDom.insertBefore(newTR, this.rows[this.currentFocusedRow]);

        this.refreshSpreadsheet();

        // reset cell's formatting and content
        for (var i=1; i<newTR.cells.length; i++)
        {
          newTR.cells[i].firstChild.innerHTML = "";
          this.unfocus(i, this.currentFocusedRow + 1);
          this.selectCell(i, this.currentFocusedRow + 1, false);
        }
      }
      catch(e)
      {
        alert(e + " - " + e.name + " - " + e.message);
      }
    },


    /**
     * Inserts a new row after the row of the current focused cell
     *
     * @type member
     * @return {void} 
     */
    insertRowAfter : function()
    {
      try
      {
        var newTR = this.rows[this.currentFocusedRow].cloneNode(true);
        ss.SpreadUtilsHtmlDom.insertAfter(newTR, this.rows[this.currentFocusedRow]);

        this.refreshSpreadsheet();

        // reset cell's formatting and content
        for (var i=1; i<newTR.cells.length; i++)
        {
          newTR.cells[i].firstChild.innerHTML = "";
          this.unfocus(i, this.currentFocusedRow + 1);
          this.selectCell(i, this.currentFocusedRow + 1, false);
        }
      }
      catch(e)
      {
        alert(e + " - " + e.name + " - " + e.message);
      }
    },


    /**
     * Remove selected rows
     *
     * @type member
     * @return {void} 
     */
    removeRows : function()
    {
      var rowsToRemove = [];

      if (this.lastSelectedRows.length > 0)
      {
        for (var i=0; i<this.lastSelectedRows.length; i++) {
          rowsToRemove[rowsToRemove.length] = this.lastSelectedRows[i];
        }
      }
      else
      {
        rowsToRemove[rowsToRemove.length] = this.currentFocusedRow;
      }

      var rowStr = "";

      for (var i=0; i<rowsToRemove.length; i++) {
        rowStr += "\t" + this.rows[rowsToRemove[i]].cells[0].firstChild.firstChild.innerHTML + "\n";
      }

      if (confirm("Are you sure you want to remove " + (rowsToRemove.length == 1 ? "this row?\n" : "these rows?\n") + rowStr))
      {
        try
        {
          for (var i=0; i<rowsToRemove.length; i++) {
            var parentNodeRmv = this.rows[rowsToRemove[i]].parentNode;
            parentNodeRmv.removeChild(this.rows[rowsToRemove[i]]);
          }

          this.refreshSpreadsheet();

          // focus on the closest upper cell
          this.focusAt(this.currentFocusedCol, this.currentFocusedRow);
        }
        catch(e)
        {
          alert(e + " - " + e.name + " - " + e.message);
        }
      }
    },


    /**
     * Remove selected cols
     *
     * @type member
     * @return {void} 
     */
    removeCols : function()
    {
      var columnsToRemove = [];

      if (this.lastSelectedColumns.length > 0)
      {
        this.debug("cas 1");
        for (var i=0; i<this.lastSelectedColumns.length; i++) {
          columnsToRemove[columnsToRemove.length] = this.lastSelectedColumns[i] + 1;
        }
      }
      else
      {
        columnsToRemove[columnsToRemove.length] = this.currentFocusedCol + 1;
      }

      var ths = this.domNode.getElementsByTagName("TH");
      var colStr = "";

      for (var i=0; i<columnsToRemove.length; i++) {
        colStr += "\t" + ths[columnsToRemove[i]].firstChild.firstChild.innerHTML + "\n";
      }

      if (confirm("Are you sure you want to remove " + (columnsToRemove.length == 1 ? "this column?\n" : "these columns?\n") + colStr))
      {
        try
        {
          for (var i=0; i<columnsToRemove.length; i++)
          {
            //qx.dom.removeNode(ths[columnsToRemove[i]]);
            var parentNodeRmv = ths[columnsToRemove[i]].parentNode;
            parentNodeRmv.removeChild(ths[columnsToRemove[i]]);

            for (var j=0; j<this.rows.length; j++) {
              //qx.dom.removeNode(this.rows[j].cells[columnsToRemove[i]]);
              var parentNodeRmv2 = this.rows[j].cells[columnsToRemove[i]].parentNode;
              parentNodeRmv2.removeChild(this.rows[j].cells[columnsToRemove[i]]);

            }
          }

          this.refreshSpreadsheet();

          // focus on the closest upper cell
          this.focusAt(this.currentFocusedCol, this.currentFocusedRow);
        }
        catch(e)
        {
          alert(e + " - " + e.name + " - " + e.message);
        }
      }
    },


    /**
     * Cut selected cells command
     *
     * @type member
     * @return {void} 
     */
    cmdCutCells : function()
    {
      this.cmdCopyCells();
      this.cmdDeleteCells();
    },


    /**
     * Copy selected cells command
     *
     * @type member
     * @return {void} 
     */
    cmdCopyCells : function()
    {
      var sel = this.getSelection();

      this.copiedCells = [];
      var j = 0;

      // sel[i][0] gives the column, sel[i][1] gives the row
      for (var i=0; i<sel.length; i++)
      {
        this.copiedCells[j] = this.getCell(sel[i][0], sel[i][1]);
        j++;
      }
    },


    /**
     * Paste on selected cells command
     *
     * @type member
     * @return {void} 
     */
    cmdPasteCells : function()
    {
      var sel = this.getSelection();

      var j = 0;

      // sel[i][0] gives the column, sel[i][1] gives the row
      for (var i=0; i<sel.length; i++)
      {
        var tdElem = this.getCell(sel[i][0], sel[i][1]);
        tdElem.formula = "" + this.copiedCells[j].formula;
        tdElem.styleList = this.copiedCells[j].styleList;
        this.evalFormula(tdElem, true);
        j++;

        if (j >= this.copiedCells.length) {
          j = 0;
        }

        this.applyStylesToCell(tdElem);
      }
    },


    /**
     * Delete selected cells command
     *
     * @type member
     * @return {void} 
     */
    cmdDeleteCells : function()
    {
      var sel = this.getSelection();

      // sel[i][0] gives the column, sel[i][1] gives the row
      for (var i=0; i<sel.length; i++)
      {
        var tdElem = this.getCell(sel[i][0], sel[i][1]);
        tdElem.formula = " ";
        this.evalFormula(tdElem, true);
      }
    },


    /**
     * Edit formula command
     *
     * @type member
     * @return {void} 
     */
    cmdEditFormula : function()
    {
      if (!this.isEditing)
      {
        var tdElem = this.getCell(this.currentFocusedCol, this.currentFocusedRow);
        this.editFormulaField(tdElem, true);
      }
    },


    /**
     * Applies one of the available formatting types to the cell (FORMATTING_TYPES.FONT, FORMATTING_TYPES.FONT_SIZE, etc)
     *
     * @type member
     * @param styleType {var} the style to change
     * @param styleValue {var} the value for the style
     * @return {void} 
     */
    formatSelectedCells : function(styleType, styleValue)
    {
      var sel = this.getSelection();

      // sel[i][0] gives the column, sel[i][1] gives the row
      for (var i=0; i<sel.length; i++)
      {
        var tdElem = this.getCell(sel[i][0], sel[i][1]);
        this.addStyleToCell(styleType, styleValue, tdElem);

        this.applyStylesToCell(tdElem);
      }
    },


    /**
     * Adds a style to the list of styles
     *
     * @type member
     * @param styleType {var} the style to change
     * @param styleValue {var} the value for the style
     * @param tdElem {var} the <TD> element to be formatted
     * @return {void} 
     */
    addStyleToCell : function(styleType, styleValue, tdElem)
    {
      if (typeof (tdElem.styleList) == "undefined") {
        tdElem.styleList = new Array();
      }

      // tdElem.styleList.add({style:styleType, value:styleValue});
      tdElem.styleList[styleType] = styleValue;
    },


    /**
     * TODOC
     *
     * @type member
     * @param style {var} TODOC
     * @param tdElem {var} TODOC
     * @return {void} 
     */
    removeStyleForCell : function(style, tdElem) {},


    /**
     * Applies the styles for a cell
     *
     * @type member
     * @param tdElem {var} the <TD> element for which styles are applied
     * @return {void} 
     */
    applyStylesToCell : function(tdElem)
    {
      var styles = tdElem.styleList;

      if (typeof (styles) != "undefined")
      {
        // font
        var font = styles[this.FORMATTING_TYPES.FONT];

        if (typeof (font) != "undefined") {
          tdElem.firstChild.style.fontFamily = font;
        }

        // font size
        var font_size = styles[this.FORMATTING_TYPES.FONT_SIZE];

        if (typeof (font_size) != "undefined") {
          tdElem.firstChild.style.fontSize = font_size;
        }

        // COLOR
        var color = styles[this.FORMATTING_TYPES.COLOR];

        if (typeof (color) != "undefined") {
          tdElem.firstChild.style.color = color;
        }

        // BG_COLOR
        var bgColor = styles[this.FORMATTING_TYPES.BG_COLOR];

        if (typeof (bgColor) != "undefined")
        {
          tdElem.firstChild.style.backgroundColor = bgColor;
          var bgColor = ss.SpreadUtilsHtmlStyle.getBackgroundColor(tdElem.firstChild);
          tdElem.firstChild.dBackgroundColor = ss.SpreadUtilsHtmlStyle.rgb2hex(bgColor[0], bgColor[1], bgColor[2]);
        }

        // BOLD
        var bold = styles[this.FORMATTING_TYPES.BOLD];

        if (typeof (bold) != "undefined")
        {
          if (bold) {
            tdElem.firstChild.style.fontWeight = "bold";
          } else {
            tdElem.firstChild.style.fontWeight = "";
          }
        }

        // ITALIC
        var italic = styles[this.FORMATTING_TYPES.ITALIC];

        if (typeof (italic) != "undefined")
        {
          if (italic) {
            tdElem.firstChild.style.fontStyle = "italic";
          } else {
            tdElem.firstChild.style.fontStyle = "";
          }
        }

        // UNDERLINE
        var underline = styles[this.FORMATTING_TYPES.UNDERLINE];

        if (typeof (underline) != "undefined")
        {
          if (underline) {
            tdElem.firstChild.style.textDecoration = "underline";
          } else {
            tdElem.firstChild.style.textDecoration = "";
          }
        }

        // Text Align
        var align = styles[this.FORMATTING_TYPES.ALIGN];

        if (typeof (align) != "undefined") {
          tdElem.firstChild.style.textAlign = align;
        }
      }
    },


    /**
     * Applies a formula
     *
     * @type member
     * @param functionName {var} TODOC
     * @return {void} 
     */
    applyFormula : function(functionName)
    {
      // minX and maxY are used to determine where the result will be written (in the bottom-left corner of the
      // most comprehensive rectangle)
      var minCol = 1000000;
      var maxRow = -1;

      var sel = this.getSelection();

      // sel[i][0] gives the column, sel[i][1] gives the row
      for (var i=0; i<sel.length; i++)
      {
        var tdElem = this.getCell(sel[i][1], sel[i][0]);
        minCol = minCol > sel[i][0] ? sel[i][0] : minCol;
        maxRow = maxRow < sel[i][1] ? sel[i][1] : maxRow;
      }

      // get the left-bottom most cell
      var resultCell = this.getCell(minCol, maxRow + 1);

      if (resultCell)
      {
        if (this.getSelectionMode() == this.SELECTION_MODES.RECTANGLE)
        {
          var cellColStart = this.fromArrToCellNotation(this.lastSelectedRegion[0], this.lastSelectedRegion[2]);
          var cellColStop = this.fromArrToCellNotation(this.lastSelectedRegion[1], this.lastSelectedRegion[3]);
          resultCell.formula = "=" + functionName + "(" + cellColStart + ":" + cellColStop + ")";
          this.evalFormula(resultCell, true);
        }
        else
        {
          // only rectangle selection mode is supported for now, so will NEVER get here
          resultCell.formula = "=" + functionName + "(";

          for (var i=0; i<sel.length; i++)
          {
            var tdElem = this.getCell(sel[i][1], sel[i][0]);
            var cellCol = String.fromCharCode("A".charCodeAt(0) + sel[i][0] - 1);
            resultCell.formula = "" + cellCol + (sel[i][1] + 1) + (i == sel.length - 1 ? "" : ",");
          }

          resultCell.formula = ")";
          this.evalFormula(resultCell, true);
        }
      }
    },


    /**
     * Recursive function for reevaluate dependent cells.
     * This method also checks for circular reference cells.
     *
     * @type member
     * @param col {var} TODOC
     * @param row {var} TODOC
     * @return {void} 
     * @throws an exception if circularities are detected
     */
    reevalDepCells : function(col, row)
    {
      // OSY DEP
      for (var i=0; i<this.depCells.length; i++)
      {
        var origCellCol = this.depCells[i].origCellCol;
        var origCellRow = this.depCells[i].origCellRow;
        var destCellCol = this.depCells[i].destCellCol;
        var destCellRow = this.depCells[i].destCellRow;

        if (destCellCol == col && destCellRow == row)
        {
          // check circular reference cells
          for (var j=0; j<this.depCells.length; j++)
          {
            var origCellColCirc = this.depCells[j].origCellCol;
            var origCellRowCirc = this.depCells[j].origCellRow;
            var destCellColCirc = this.depCells[j].destCellCol;
            var destCellRowCirc = this.depCells[j].destCellRow;

            if (origCellColCirc == destCellCol && origCellRowCirc == destCellRow && destCellColCirc == origCellCol && destCellRowCirc == origCellRow) {
              throw "Circular Reference detected for this cell";
            }
          }

          var cell = this.getCell(origCellCol, origCellRow);
          this.evalFormula(cell, true, true);

          // recursive call to reeval dependent cells tree
          this.reevalDepCells(origCellCol, origCellRow);
        }
      }
    },


    /**
     * Returns an array containing (col,row) as they are retrieved from the cell notation (A2, B33, etc).
     * The indexes are relative to the cells spreadsheet (col 0 - is actually the second column - 1st one is the header)
     *
     * @type member
     * @param cellNotation {var} a string representing the cell notation
     * @return {Map} an object with two properties: col and row
     */
    fromCellNotationToArr : function(cellNotation)
    {
      var letter = "" + cellNotation.match(/[a-zA-Z]*/);
      var number = "" + cellNotation.match(/[0-9]+/);

      // build the indices for the cell
      var col = this.fromCharToIdx(letter);
      var row = parseInt(number) - 1;

      return {
        "col" : col,
        "row" : row
      };
    },


    /**
     * Creates a cell notation, given a cell and a row. The cell and row must start from 0 (this means, you can't pass
     * rowIndex and cellIndex properties of <TR> and <TD> elements. You need to handle conversion from these properties
     * to proper indexes (usually this is transparently handled by getCellRow and getCellCol). A (0,0) will be A1 in cell
     * notation
     *
     * @type member
     * @param col {var} the column index (for 0 a value of 'A' will be returned)
     * @param row {var} the row index (for 0 a value of 1 will be returned)
     * @return {var} a string representing the cell notation A1, C5, etc
     */
    fromArrToCellNotation : function(col, row) {
      return this.fromIdxToChar(col) + (row + 1);
    },


    /**
     * Utility function for transforming a char code to it's correspondent in number (a-0, b-1,...)
     *
     * @type member
     * @param ch {var} the char
     * @return {var} a number
     */
    fromCharToIdx : function(ch) {
      return ch.toLowerCase().charCodeAt(0) - "a".charCodeAt(0);
    },


    /**
     * Utility function for transforming and index to its corresponding char (0-a, 1-b, ...)
     *
     * @type member
     * @param idx {var} TODOC
     * @return {var} an uppercase string with one letter
     */
    fromIdxToChar : function(idx) {
      return String.fromCharCode("a".charCodeAt(0) + idx).toUpperCase();
    },


    /**
     * Retrieves the <TD> element based on the cell notation (A1, C5, etc)
     *
     * @type member
     * @param cellStr {var} a string representing the cell notation
     * @return {var} the <TD> object
     */
    getCellByCellNotation : function(cellStr)
    {
      var cellCoords = this.fromCellNotationToArr(cellStr);

      return this.getCell(cellCoords.col, cellCoords.row);
    },


    /**
     * Returns the row for a cell in the spreadsheet given a <TD> element. Note IE keeps track of TH elements also.
     * (ignoring cell headers)
     *
     * @type member
     * @param tdElem {var} the <TD> element
     * @return {var} the row index
     */
    getCellRow : function(tdElem) {
      return qx.core.Client.getInstance().isMshtml() ? tdElem.parentNode.rowIndex - 1 : tdElem.parentNode.rowIndex;
    },


    /**
     * Returns the column (within the spreadsheet) for a cell in the spreadsheet for a <TD> element (ignoring cell headers)
     *
     * @type member
     * @param tdElem {var} the <TD> element
     * @return {var} the column index
     */
    getCellCol : function(tdElem) {
      return tdElem.cellIndex - 1;
    },


    /**
     * Returns the <TD> element for a given cell in the spreadsheet. For example, getCell(0,0) should return the first cell
     * on the first row (ignoring first column and the thead).
     *
     * @type member
     * @param col {var} a number representing the col of the cell in the spreadsheet. Starts from 0 (0 being the equivalent of column A).
     * @param row {var} a number representing the row of the cell in the spreadsheet. Starts from 0 (0 being the equivalent of row 1).
     * @return {var} a <TD> object
     */
    getCell : function(col, row) {
      return this.rows[row].cells[col + 1];
    },


    /**
     * Returns the number of quotation marks ignoring the escaped one
     *
     * @type member
     * @param str {String} the string too look for quotation inside
     * @param start {var} an integer representing the position to start looking for
     * @param stop {var} an integer representing the position to stop looking for
     * @return {var} TODOC
     */
    getNumberOfQuotes : function(str, start, stop)
    {
      var numQuotes = 0;

      for (var i=start; i<stop; i++)
      {
        if (str.charAt(i) == '"')
        {
          if (numQuotes % 2 == 1 && i > 0 && str.charAt(i - 1) == "\\")
          {
            if (i > 1 && str.charAt(i - 2) == "\\") {
              numQuotes++;
            }
          }
          else
          {
            numQuotes++;
          }
        }
      }

      return numQuotes;
    },


    /**
     * Evaluates a formula for a cell, and puts the result back in the cell
     *
     * @type member
     * @param _cell {var} TODOC
     * @param evaluate {var} TODOC
     * @param noDependencies {var} TODOC
     * @return {string | var} the result of the formula
     * @throws TODOC
     */
    evalFormula : function(_cell, evaluate, noDependencies)
    {
      var cellNotation = this.fromArrToCellNotation(this.getCellCol(_cell), this.getCellRow(_cell));
      var formula = _cell.formula;
      var errors = false;

      if (formula == "" || typeof (formula) == "undefined") {
        return "";
      }


      // a cell's value it's a formula only if it starts with "="
      if (!evaluate && _cell.cellType != this.CELL_TYPES.FORMULA) {
        return _cell.formula;
      }


      /*
       * movie (flash object)
       */
      var startUrlPos = 0;
      this.debug("=====> _cell.formula : " + _cell.formula);
      if (_cell.formula.substr(0, 13).toLowerCase() == "=flashmovie(\"") {
         startUrlPos = 13;
         var lengthUrl = _cell.formula.indexOf("\")") - startUrlPos;
         var url = _cell.formula.substr(startUrlPos, lengthUrl);
         _cell.firstChild.innerHTML = "<object type=\"application/x-shockwave-flash\" data=\"" + url + "\" width=\"425\" height=\"350\">" +
                                    "<param name=\"movie\" value=\"" + url + "\" />" +
                                    "<param name=\"wmode\" value=\"transparent\" />" +
                                    "</object>";
         _cell.cellType = this.CELL_TYPES.STRING;
         return _cell.firstChild.innerHTML;

      }
      /*
       * Web Page
       */
      else if (_cell.formula.substr(0, 10).toLowerCase() == "=webpage(\"") {
        startUrlPos = 10;
         var lengthUrl = _cell.formula.indexOf("\")") - startUrlPos;
         var url = _cell.formula.substr(startUrlPos, lengthUrl);
         _cell.firstChild.innerHTML = "<object type=\"text/html\" data=\"" + url + "\" width=\"425\" height=\"350\">" +
                                    "</object>";
         _cell.cellType = this.CELL_TYPES.STRING;
         return _cell.firstChild.innerHTML;

      }
      /*
       * PDF File
       */
      else if (_cell.formula.substr(0, 10).toLowerCase() == "=pdffile(\"") {
        startUrlPos = 10;
         var lengthUrl = _cell.formula.indexOf("\")") - startUrlPos;
         var url = _cell.formula.substr(startUrlPos, lengthUrl);
         _cell.firstChild.innerHTML = "<object type=\"application/pdf\" data=\"" + url + "\" width=\"425\" height=\"350\">" +
                                    "</object>";
         _cell.cellType = this.CELL_TYPES.STRING;
         return _cell.firstChild.innerHTML;

      }
      /*
       * Word File
       */
      else if (_cell.formula.substr(0, 11).toLowerCase() == "=wordfile(\"") {
        startUrlPos = 11;
         var lengthUrl = _cell.formula.indexOf("\")") - startUrlPos;
         var url = _cell.formula.substr(startUrlPos, lengthUrl);
         _cell.firstChild.innerHTML = "<object type=\"application/msword\" data=\"" + url + "\" width=\"425\" height=\"350\">" +
                                    "</object>";
         _cell.cellType = this.CELL_TYPES.STRING;
         return _cell.firstChild.innerHTML;

      }
      /*
       * Excel File
       */
      else if (_cell.formula.substr(0, 12).toLowerCase() == "=excelfile(\"") {
        startUrlPos = 12;
         var lengthUrl = _cell.formula.indexOf("\")") - startUrlPos;
         var url = _cell.formula.substr(startUrlPos, lengthUrl);
         _cell.firstChild.innerHTML = "<object type=\"application/msexcel\" data=\"" + url + "\" width=\"425\" height=\"350\">" +
                                    "</object>";
         _cell.cellType = this.CELL_TYPES.STRING;
         return _cell.firstChild.innerHTML;

      }



      var result;

      if (trim("" + formula).charAt(0) != "=")
      {
        // try to detect a type for the cell (date, string, number)
        var tmp = formula.toLowerCase();

        // check if it's a date
        var _val;

        if ((_val = this.parseDateFormat1(tmp)) != null) {
          _cell.cellType = this.CELL_TYPES.DATE;
        } else if ((_val = this.parseDateFormat2(tmp)) != null) {
          _cell.cellType = this.CELL_TYPES.DATE;
        }
        else if ((_val = this.parseNumberFormat1(tmp)) != null)
        {
          // check if it's a number
          _cell.cellType = this.CELL_TYPES.NUMBER;
        }
        else
        {
          // if none of the above, consider it a string
          _val = formula;
          _cell.cellType = this.CELL_TYPES.STRING;
        }

        _cell.formula = _val;
        result = _val;
      }
      else
      {
        _cell.cellType = this.CELL_TYPES.FORMULA;

        this.debug("STEP 1: transform intervals of cells for " + cellNotation + ": " + _cell.formula);


        /********************************************************************
        STEP 1: transform intervals of cells to array of cells (B2:C3 => B2,B3,C2,C3
        *********************************************************************/
        var pcs = formula.match(this.reInterval);

        var inlinedFormula = formula;

        if (pcs != null)
        {
          // for each interval
          // TODO check the cells is not part of a string
          for (var i=0; i<pcs.length; i++)
          {
            var cells = pcs[i].split(":");

            var cell0 = this.fromCellNotationToArr(cells[0]);
            var cell1 = this.fromCellNotationToArr(cells[1]);

            // build a string with all the intermediary values
            var inlinedInterval = "";

            for (var colIdx=cell0.col; colIdx<=cell1.col; colIdx++)
            {
              for (var rowIdx=cell0.row; rowIdx<=cell1.row; rowIdx++) {
                inlinedInterval += this.fromIdxToChar(colIdx) + (rowIdx + 1) + ",";
              }
            }

            // remove the last comma
            inlinedInterval = inlinedInterval.substring(0, inlinedInterval.length - 1);

            // replaced interval with inline formula
            // for each occ of cells, replace it with inlinedInterval
            var cellsPos = 0;

            while ((cellsPos = inlinedFormula.indexOf(pcs[i], cellsPos)) != -1)
            {
              // check the number of " chars before pcs[i]
              var numQuotes = this.getNumberOfQuotes(inlinedFormula, 0, cellsPos);

              if (numQuotes % 2 == 0) {
                inlinedFormula = inlinedFormula.substring(0, cellsPos) + inlinedInterval + inlinedFormula.substring(cellsPos + pcs[i].length);
              }

              cellsPos += pcs[i].length;
            }
          }
        }

        this.debug("STEP 2: search for cell notations and mark the graph for " + cellNotation + ": " + inlinedFormula);


        /********************************************************************
        STEP 2: search for cell notations that are not inside strings and mark the graph
        *********************************************************************/
        // reset the dependencies for this cell
        // OSY DEP
        for (var i=0; i<this.depCells.length; i++)
        {
          if (this.depCells[i].origCellCol == this.getCellCol(_cell) && this.depCells[i].origCellRow == this.getCellRow(_cell)) {
            qx.lang.Array.removeAt(this.depCells, i);
          }
        }

        var cellsMatches = inlinedFormula.match(this.reCell);

        if (cellsMatches != null)
        {
          // for each cell
          for (var i=0; i<cellsMatches.length; i++)
          {
            var cell = cellsMatches[i].substring(1);

            var cellPos = 0;

            while ((cellPos = inlinedFormula.indexOf(cell, cellPos)) != -1)
            {
              // check the number of " chars before pcs[i]
              var numQuotes = this.getNumberOfQuotes(inlinedFormula, 0, cellPos);

              if (numQuotes % 2 == 0)
              {
                // mark dependent cells
                var cellCoords = this.fromCellNotationToArr(cell);

                // OSY DEP
                var curIdx = this.depCells.length;
                this.depCells[curIdx] = new Object();
                this.depCells[curIdx].origCellCol = this.getCellCol(_cell);
                this.depCells[curIdx].origCellRow = this.getCellRow(_cell);
                this.depCells[curIdx].destCellCol = cellCoords.col;
                this.depCells[curIdx].destCellRow = cellCoords.row;
              }

              cellPos += cell.length;
            }
          }
        }

        try
        {
          this.debug("STEP 3: check circularities for " + cellNotation + ", formula: " + inlinedFormula);


          /********************************************************************
           STEP 3: check circularities
          *********************************************************************/
          // this.checkCircularitiesFromNode(originalGraphNode);
          this.debug("STEP 4: evaluate each cell's value for " + cellNotation);


          /********************************************************************
           STEP 4: if no circularities, replace each cell name with its value
          *********************************************************************/
          var cells = inlinedFormula.match(this.reCell);

          if (cells != null)
          {
            // for each cell
            // TODO check the cells is not part of a string
            for (var i=0; i<cells.length; i++)
            {
              var cell = cells[i].substring(1);

              // replaced cell with it.lengts value
              var cellPos = 0;

              while ((cellPos = inlinedFormula.indexOf(cell, cellPos)) != -1)
              {
                // check the number of " chars before pcs[i]
                var numQuotes = this.getNumberOfQuotes(inlinedFormula, 0, cellPos);

                if (numQuotes % 2 == 0)
                {
                  var cellValue;
                  var cellObj = this.getCellByCellNotation(cell);

                  try {
                    cellValue = this.evalFormula(cellObj, false, true);
                  } catch(e) {
                    throw e;
                  }

                  // for strings, the value to replace has to contain the quotes
                  var toReplace = "";

                  if (cellObj.cellType == this.CELL_TYPES.NUMBER) {
                    toReplace = cellValue;
                  }

                  if (cellObj.cellType == this.CELL_TYPES.STRING) {
                    toReplace = "\"" + cellValue + "\"";
                  }

                  if (cellObj.cellType == this.CELL_TYPES.FORMULA) {
                    toReplace = cellValue;
                  }

                  // if cell has no value, remove previous , if exists or if previous char is ( remove next ,
                  // make sure the cell that have no value don't affect the formula by leaving random commas
                  if (toReplace == "")
                  {
                    var tmp = trim(inlinedFormula.substring(0, cellPos));
                    var tmp2 = trim(inlinedFormula.substring(cellPos + cell.length));

                    // two cases: comma is before and after
                    if (cellPos > 0 && tmp.charAt(tmp.length - 1) == ',') {
                      inlinedFormula = inlinedFormula.substring(0, cellPos - 1) + toReplace + inlinedFormula.substring(cellPos + cell.length);
                    }
                    else
                    {
                      if (cellPos > 0 && tmp.charAt(tmp.length - 1) == '(' && cellPos < inlinedFormula.length - 1 && tmp2.charAt(0) == ',') {
                        inlinedFormula = tmp + toReplace + tmp2.substring(1);
                      }
                    }
                  }
                  else
                  {
                    inlinedFormula = inlinedFormula.substring(0, cellPos) + toReplace + inlinedFormula.substring(cellPos + cell.length);
                  }
                }

                cellPos += cell.length;
              }
            }
          }

          this.debug("STEP 5: turn functions to lower case for " + cellNotation + ", formula: " + inlinedFormula);


          /********************************************************************
          STEP 5: turn function to lower case
          *********************************************************************/
          var functions = inlinedFormula.match(this.reFunction);

          if (functions != null)
          {
            // for each fn
            for (var i=0; i<functions.length; i++)
            {
              var fnToLower = functions[i].toLowerCase();

              // only if it's not lower already
              if (fnToLower != functions[i])
              {
                var fnPos = 0;

                while ((fnPos = inlinedFormula.indexOf(functions[i], fnPos)) != -1)
                {
                  // check the number of " chars before pcs[i]
                  var numQuotes = this.getNumberOfQuotes(inlinedFormula, 0, fnPos);

                  if (numQuotes % 2 == 0) {
                    inlinedFormula = inlinedFormula.substring(0, fnPos) + fnToLower + inlinedFormula.substring(fnPos + functions[i].length);
                  }

                  fnPos += functions[i].length;
                }
              }
            }
          }

          // and now evaluate the formula
          result = eval(inlinedFormula.substring(1));

          this.debug("result formula for " + cellNotation + ", formula: " + inlinedFormula);
        }
        catch(e)
        {
          this.debug(e);

          // throw(e);
          result = e;
          errors = true;
        }
      }

      if (!errors)
      {
        if (!noDependencies)
        {
          this.debug("STEP 6: re-evaluate cells dependent on this very cell: " + cellNotation + ", formula: " + inlinedFormula);


          /********************************************************************
          STEP 6: reevaluate dependent cells
          *********************************************************************/
          // OSY DEP
          this.reevalDepCells(this.getCellCol(_cell), this.getCellRow(_cell));
        }
        else
        {
          this.debug("STEP 6: dependencies will NOT be re-evaluated for cell: " + cellNotation);
        }
      }

      // Align Right if numeric
      var resultStr = "" + result;
      var _num = resultStr.match(/[0-9e.,-]+/);

      if (_num != null && _num == resultStr) {
        _cell.firstChild.style.textAlign = "right";
      }

      // replace dot by the decimal separator
      if (typeof (result) == "number")
      {
        result = "" + result;
        result = result.replace(/[.]+/, this.DECIMAL_SEP);
      }

      _cell.firstChild.innerHTML = result;
      





      this.debug("Final result for " + cellNotation + ": " + result);

      return result;  // return this.formatCellForPresentation(cell);;
    },


    /**
     * Fires an event
     *
     * @type member
     * @param evt {Event} the event to be fired
     * @return {void} 
     */
    _fireEvent : function(evt)
    {
      if (typeof this[evt] == "function")
      {
        var args = [ this ];

        for (var i=1; i<arguments.length; i++) {
          args.push(arguments[i]);
        }

        this[evt].apply(this, args);
      }
    },


    /**
     * Formats the cell for presentation. The formatting should be based on cell's type (number, date, string) and cell's
     * format. Possible formats for date are "MM/DD/YYYY", "DD-MMM-YYYY", etc.
     * Currently cell's formats are not supported yet.
     *
     * @type member
     * @param cell {var} the <TD> element to be formatted for presentation
     * @return {var} the formatted value
     */
    formatCellForPresentation : function(cell)
    {
      var formula = cell.formula;

      if (formula && typeof (formula) != "undefined" && formula != null)
      {
        // now, the cell format should be checked...
        if (formula instanceof Date)
        {
          with (formula) {
            return getDate() + "-" + this.months[getMonth()].toUpperCase() + "-" + getFullYear();
          }
        }

        // now, the cell format should be checked...
        if (typeof (formula) == "number") {
          return formula;
        }
      }
    },


    /**
     * Creates a date from a string parsing format 1: MM/DD/YYYY
     *
     * @type member
     * @param str {String} the string to be parsed
     * @return {var | null} a Date object if parsing succeeded or null if string could not be parsed
     */
    parseDateFormat1 : function(str)
    {
      var _date = str.match(this.reDate1);

      if (_date && _date != null)
      {
        _date = "" + _date;

        if (_date == trim(str))
        {
          var vals = _date.split("/");
          var month = parseInt(vals[0]);
          var day = parseInt(vals[1]);
          var year = parseInt(vals[2]);

          if (this.isDateValid(day, month, year))
          {
            if (year < 100 && year >= 30) {
              year = 2000 + year;
            } else if (year < 100) {
              year = 1900 + year;
            }

            var dateObj = new Date();
            dateObj.setFullYear(year);
            dateObj.setMonth(month - 1);
            dateObj.setUTCDate(day);

            return dateObj;
          }
        }
      }

      return null;
    },


    /**
     * Creates a date from a string parsing format 2: DD-MMM-YYYY
     *
     * @type member
     * @param str {String} the string to be parsed
     * @return {null | var} a Date object if parsing succeeded or null if string could not be parsed
     */
    parseDateFormat2 : function(str)
    {
      var _date = str.match(this.reDate2);

      if (_date && _date != null)
      {
        _date = "" + _date;

        if (_date == trim(str))
        {
          var vals = _date.split("-");

          var month = vals[1];
          var found = -1;

          for (var i=0; i<this.months.length; i++)
          {
            if (this.months[i] == month)
            {
              found = i;
              break;
            }
          }

          if (found == -1) {
            return null;
          }

          var day = parseInt(vals[0]);
          var year = parseInt(vals[2]);

          if (this.isDateValid(day, month, year))
          {
            if (year < 100 && year >= 30) {
              year = 2000 + year;
            } else if (year < 100) {
              year = 1900 + year;
            }

            var dateObj = new Date();
            dateObj.setFullYear(year);
            dateObj.setMonth(found);
            dateObj.setUTCDate(day);

            return dateObj;
          }
        }
      }

      return null;
    },


    /**
     * Creates a number from a string
     *
     * @type member
     * @param str {String} the string to be parsed
     * @return {null | var} an int or a float if parsing succeeded or null if string could not be parsed
     */
    parseNumberFormat1 : function(str)
    {
      var _num = str.match(this.reNumber1);

      if (_num && _num != null) {
        return null;
      }

      var dotIdx = str.indexOf(".");

      if (dotIdx != -1 && dotIdx != str.lastIndexOf(".")) {
        return null;
      }

      var comIdx = str.indexOf(",");

      if (comIdx != -1)
      {
        if (comIdx != str.lastIndexOf(",")) {
          return null;
        }

        str = str.replace(/,/, ".");
      }

      var eIdx = str.indexOf("e");

      if (eIdx != -1 && eIdx != str.lastIndexOf("e")) {
        return null;
      }

      if (dotIdx != -1 && eIdx != -1 && dotIdx + 1 != eIdx) {
        return null;
      }

      var _int = parseInt(str);

      if ("" + _int == str) {
        return _int;
      }

      var _float = parseFloat(str);

      if (!isNaN(_float)) {
        return _float;
      }

      return null;
    },


    /**
     * Validates a date. Not implemented yet - just a simple check the ranges for each param are good
     *
     * @type member
     * @param day {var} TODOC
     * @param month {var} TODOC
     * @param year {var} TODOC
     * @return {boolean} TODOC
     */
    isDateValid : function(day, month, year)
    {
      if (month < 1 && month > 12 && day < 1 && day > 31) {
        return false;
      }

      // TODO check valid date
      return true;
    },


    /**
     * Spreadsheet specific debug function
     *
     * @type member
     * @param msg {var} a string representing the message to debug
     * @return {void} 
     */
    debugSpread : function(msg)
    {
      var date = new Date();
      var tmp = date.getHours() + ":" + date.getMinutes() + "-" + date.getSeconds() + ":" + date.getMilliseconds() + " :  " + msg;
      var dbg = document.getElementById("debugDiv");
      dbg.value = tmp + "\n" + dbg.value;
      this.debug(tmp);
    },


    /**
     * Adds "px" to a number
     *
     * @type member
     * @param num {Number} the number
     * @return {var} TODOC
     */
    toPx : function(num) {
      return num + "px";
    }
  }
});


/********************************************************************************
* MATHEMATICAL FUNCTIONS
*********************************************************************************/
function sum()
{
  var _sum = 0;

  if (arguments != null && arguments.length > 0)
  {
    for (var i=0; i<arguments.length; i++) {
      _sum += arguments[i];
    }
  }

  return _sum;
}

function avg()
{
  var _avg = 0;

  if (arguments != null && arguments.length > 0)
  {
    for (var i=0; i<arguments.length; i++) {
      _avg += arguments[i];
    }

    _avg /= arguments.length;
  }

  return _avg;
}

function min()
{
  var _min = Number.POSITIVE_INFINITY;

  if (arguments != null && arguments.length > 0)
  {
    for (var i=0; i<arguments.length; i++)
    {
      if (_min > arguments[i]) {
        _min = arguments[i];
      }
    }
  }

  return _min;
}

function max()
{
  var _max = Number.NEGATIVE_INFINITY;

  if (arguments != null && arguments.length > 0)
  {
    for (var i=0; i<arguments.length; i++)
    {
      if (_max < arguments[i]) {
        _max = arguments[i];
      }
    }
  }

  return _max;
}

function count()
{
  var _count = 0;

  if (arguments != null && arguments.length > 0)
  {
    for (var i=0; i<arguments.length; i++)
    {
      if (arguments[i] != null && typeof (arguments[i]) != "undefined" && arguments[i] != "") {
        _count++;
      }
    }
  }

  return _count;
}

function product()
{
  var _product = 1;

  if (arguments != null && arguments.length > 0)
  {
    for (var i=0; i<arguments.length; i++) {
      _product *= arguments[i];
    }
  }

  return _product;
}

var abs = Math.abs;
var acos = Math.acos;
var asin = Math.asin;
var atan = Math.atan;
var atan2 = Math.atan2;
var ceil = Math.ceil;
var cos = Math.cos;
var exp = Math.exp;
var floor = Math.floor;
var log = Math.log;
var pow = Math.pow;
var random = Math.random;
var round = Math.round;
var sin = Math.sin;
var sqrt = Math.sqrt;
var tan = Math.tan;


/********************************************************************************
* TEXT FUNCTIONS
*********************************************************************************/
function len()
{
  if (arguments != null && arguments.length > 0) {
    return arguments[0].length;
  }
}

function lower()
{
  if (arguments != null && arguments.length > 0) {
    return arguments[0].toLowerCase();
  }
}

function upper()
{
  if (arguments != null && arguments.length > 0) {
    return arguments[0].toUpperCase();
  }
}

function left()
{
  if (arguments != null && arguments.length > 0)
  {
    var endPos = 1;

    if (arguments[1]) {
      endPos = arguments[1];
    }

    return arguments[0].substring(0, endPos);
  }
}

function right()
{
  if (arguments != null && arguments.length > 0)
  {
    var startPos = 1;

    if (arguments[1]) {
      startPos = arguments[1];
    }

    return arguments[0].substring(arguments[0].length - startPos);
  }
}

var trim = qx.lang.String.trim;
