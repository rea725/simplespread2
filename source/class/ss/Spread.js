
/**
 * SpreadSheet project
 * Author : O. Sadey
 */
qx.Class.define("qx.ui.spread.Spread",
{
  extend : qx.ui.layout.VerticalBoxLayout,




  /*
  *****************************************************************************
     CONSTRUCTOR
  *****************************************************************************
  */

  construct : function()
  {
    qx.ui.layout.VerticalBoxLayout.call(this);

    this.auto();

    this.setMinWidth(200);
    this.setMinHeight(200);

    this.version = "X4VIEW Web SpreadSheet component v0.1.0";
    this.author = "Author : O. SADEY";

    this.cliDoc = qx.ui.core.ClientDocument.getInstance();

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

    this.formulaField = "";

    this.prevThis = this;

    this.tc = null;
    this.tabContainer = null;
    this.tabView = null;
    this.tabViewPages = [];
    this.tabViewButtons = [];
    this.tabViewFormula = [];
    this.tabs = [];
    this.activeTab = null;

    this.styleBold = null;
    this.styleItalic = null;
    this.styleUnderline = null;

    this.boldItem = null;
    this.italicItem = null;
    this.underlineItem = null;
    this.colorItem = null;
    this.bgcolorItem = null;
    this.fontMenu = null;
    this.fontSizeMenu = null;
    this.rowsMenu = null;
    this.colsMenu = null;
    this.sheetMenu = null;
    this.functionsMenu = null;

    // osy
    this.createSpread();
  },



  /*
  *****************************************************************************
     EVENTS
  *****************************************************************************
  */

  events: {

    /** Fired when a new document was selected.*/
    "new"     : "qx.event.type.DataEvent",
    /** Fired when a document was opened.*/
    "open"     : "qx.event.type.DataEvent",
    /** Fired when a document was saved. The event holds the new save data in its data property.*/
    "save"     : "qx.event.type.DataEvent",
    /** Fired when a document was saved (save as...). The event holds the new selected date in its data property.*/
    "saveas"     : "qx.event.type.DataEvent",
    /** Fired when a document was closed.*/
    "close"     : "qx.event.type.DataEvent",
    /** Fired when a sheet was closed.*/
    "closesheet"     : "qx.event.type.DataEvent"

  },





  /*
  *****************************************************************************
     MEMBERS
  *****************************************************************************
  */

  members :
  {
    
    /**
     * Creates a new sheet
     *
     * @type member
     * @param sheetName {var} TODOC
     * @return {void} 
     */
    createSheet : function(sheetName)
    {
        this.initSheet(this.tabs.length, sheetName, "", "");

    },

    /**
     * Open an existing spreadsheet with serialized XML Data
     *
     * @spreadData spread data (xml)
     * @spreadNamePrefix sheet name prefix
     * @return {void} 
     */
    openSpread : function(spreadData, sheetNamePrefix)
    {
        var sprData = spreadData.split("**sheetdata**");
        for (var i=0; i<sprData.length - 1; i++) {
            var sheetName = sheetNamePrefix + " " + (i+1);
            this.debug("======> sprData[i] : " + sprData[i]);
            var spreadData2 = sprData[i].split("**additionaldomdata**");
            this.initSheet(i, sheetName, spreadData2[0], spreadData2[1]);

        }
    },



    /**
     * Create toolbar, tabcontainer and first sheet
     *
     * @type member
     * @return {void} 
     */
    createSpread : function()
    {
      this.createToolbar();

      this.createTabContainer();

      //this.createSheet("sheet 1");
    },

    //    this.createSheet("sheet 2");
    //    this.createSheet("sheet 3");
    /**
     * TODOC
     *
     * @type member
     * @return {void} 
     */
    createToolbar : function()
    {
      /*
       * Commands (Menubar shortcuts)
       */

      var prevThis = this;

      this.cdeMenu = new qx.client.Command();

      /* All other commands (copy, cut...) */
      this.cdeMenu.addEventListener("execute", function(e) {
        this.debug("Execute: " + e.getData().getLabel());

      });

      this.cdeNew = new qx.client.Command("Ctrl+N");

      this.cdeNew.addEventListener("execute", function(e) {
        this.debug("Execute: New");
        prevThis.createDispatchDataEvent("new", "data...");
      });

      this.cdeOpen = new qx.client.Command("Ctrl+O");

      this.cdeOpen.addEventListener("execute", function(e) {
        this.debug("Execute: Open");
        prevThis.createDispatchDataEvent("open", "data...");
      });

      this.cdeSave = new qx.client.Command("Ctrl+S");

      this.cdeSave.addEventListener("execute", function(e) {
        this.debug("Execute: Save");
        var strBuilder = new qx.util.StringBuilder();
        var saveData = "";
        for (var i=0; i<prevThis.tabs.length; i++) {
            var elem = prevThis.tabs[i].getElement();
            var serial = qx.xml.Element.serialize(elem.firstChild);

            // Get additional Dom Data
            var additionalData = prevThis.tabs[i].sheet.getSerialAdditData();

            strBuilder.add(serial + "**additionaldomdata**" + additionalData + "**sheetdata**");
        }
        saveData = strBuilder.get();
        prevThis.createDispatchDataEvent("save", saveData);
      });

      this.cdeSaveAs = new qx.client.Command();

      this.cdeSaveAs.addEventListener("execute", function(e) {
        this.debug("Execute: SaveAs");
        prevThis.createDispatchDataEvent("saveas", "data...");
      });
      
      this.cdeClose = new qx.client.Command();

      this.cdeClose.addEventListener("execute", function(e) {
        this.debug("Execute: Close");
        prevThis.createDispatchDataEvent("close", "data...");
      });


      this.cdeUndo = new qx.client.Command("Ctrl+Z");

      this.cdeUndo.addEventListener("execute", function(e) {
        this.debug("Execute: Undo");
      });

      this.cdeRedo = new qx.client.Command("Ctrl+Y");

      this.cdeRedo.addEventListener("execute", function(e) {
        this.debug("Execute: Redo");
      });

      this.cdeCut = new qx.client.Command("Ctrl+X");

      this.cdeCut.addEventListener("execute", function(e)
      {
        this.debug("Execute: Cut");
        prevThis.cmdCut();
      });

      this.cdeCopy = new qx.client.Command("Ctrl+C");

      this.cdeCopy.addEventListener("execute", function(e)
      {
        this.debug("Execute: Copy");
        prevThis.cmdCopy();
      });

      this.cdePast = new qx.client.Command("Ctrl+V");

      this.cdePast.addEventListener("execute", function(e)
      {
        this.debug("Execute: Past");
        prevThis.cmdPaste();
      });

      this.cdeDelCells = new qx.client.Command("Delete");

      this.cdeDelCells.addEventListener("execute", function(e)
      {
        this.debug("Execute: Del Cells");
        prevThis.cmdDelCells();
      });
      
      this.cdeDelRows = new qx.client.Command();

      this.cdeDelRows.addEventListener("execute", function(e)
      {
        this.debug("Execute: Del Selected Rows");
        prevThis.cmdDelRows();
      });

      this.cdeDelColumns = new qx.client.Command();

      this.cdeDelColumns.addEventListener("execute", function(e)
      {
        this.debug("Execute: Del Selected Columns");
        prevThis.cmdDelColumns();
      });


      this.cdeInsertRow = new qx.client.Command();

      this.cdeInsertRow.addEventListener("execute", function(e)
      {
        this.debug("Execute: Insert Row");
        prevThis.cmdInsertRow();
      });

      this.cdeInsertColumn = new qx.client.Command();

      this.cdeInsertColumn.addEventListener("execute", function(e)
      {
        this.debug("Execute: Insert Column");
        prevThis.cmdInsertColumn();
      });

      this.cdeSelectAll = new qx.client.Command("Ctrl+A");

      this.cdeSelectAll.addEventListener("execute", function(e) {
        this.debug("Execute: Select All");
      });

      this.cdeFind = new qx.client.Command("Ctrl+F");

      this.cdeFind.addEventListener("execute", function(e) {
        this.debug("Execute: Find");
      });

      this.cdeFindAgain = new qx.client.Command("F3");

      this.cdeFindAgain.addEventListener("execute", function(e) {
        this.debug("Execute: Find Again");
      });

      this.cdeEditFormula = new qx.client.Command("F2");

      this.cdeEditFormula.addEventListener("execute", function(e)
      {
        this.debug("Execute: Edit Formula");
        prevThis.cmdEditFormula();
      });
      
      this.cdeAbout = new qx.client.Command();
      this.cdeAbout.addEventListener("execute", function(e)
      {
        this.debug("Execute: About");
        prevThis.cmdAbout();
      });

      /*
       * MenuBar
       */

      var m1 = new qx.ui.menu.Menu;

      var mb1_01 = new qx.ui.menu.Button("New", null, this.cdeNew);
      var mb1_02 = new qx.ui.menu.Button("Open...", null, this.cdeOpen);
      var mb1_03 = new qx.ui.menu.Button("Save", null, this.cdeSave);
      var mb1_04 = new qx.ui.menu.Button("Save as...", null, this.cdeSaveAs);
      var mb1_05 = new qx.ui.menu.Button("Close", null, this.cdeClose);

      m1.add(mb1_01, mb1_02, mb1_03, mb1_04, mb1_05);

      var m2 = new qx.ui.menu.Menu;

      var mb2_01 = new qx.ui.menu.Button("Undo", null, this.cdeUndo);
      var mb2_02 = new qx.ui.menu.Button("Redo", null, this.cdeRedo);
      var mb2_b1 = new qx.ui.menu.Separator();
      var mb2_03 = new qx.ui.menu.Button("Cut", "icon/16/actions/edit-cut.png", this.cdeCut);
      var mb2_04 = new qx.ui.menu.Button("Copy", "icon/16/actions/edit-copy.png", this.cdeCopy);
      var mb2_05 = new qx.ui.menu.Button("Paste", "icon/16/actions/edit-paste.png", this.cdePast);
      var mb2_06 = new qx.ui.menu.Button("Delete", "icon/16/actions/edit-delete.png", this.cdeDelCells);
      var mb2_b2 = new qx.ui.menu.Separator();
      var mb2_07 = new qx.ui.menu.Button("Select All", null, this.cdeSelectAll);
      var mb2_08 = new qx.ui.menu.Button("Find", null, this.cdeFind);
      var mb2_09 = new qx.ui.menu.Button("Find Again", null, this.cdeFindAgain);

      // mb2_05.setEnabled(false);
      // mb2_06.setEnabled(false);
      mb2_09.setEnabled(false);

      m2.add(mb2_01, mb2_02, mb2_b1, mb2_03, mb2_04, mb2_05, mb2_06, mb2_b2, mb2_07, mb2_08, mb2_09);

      var m3 = new qx.ui.menu.Menu;

      var mb3_suba_01 = new qx.ui.menu.Button("About...", "icon/16/actions/help-about.png", this.cdeAbout);

      m3.add(mb3_suba_01);

      this.cliDoc.add(m1, m2, m3);

      var mb1 = new qx.ui.toolbar.ToolBar;
      mb1.setWidth("100%");

      //var mp1 = new qx.ui.toolbar.Part;

      //mb1.add(mp1);

      var mbb1 = new qx.ui.toolbar.MenuButton("File", m1);//, "icon/16/actions/document-new.png");
      var mbb2 = new qx.ui.toolbar.MenuButton("Edit", m2);//, "icon/16/actions/edit.png");
      var mbb3 = new qx.ui.toolbar.MenuButton("?", m3);//, "icon/16/apps/preferences-desktop-wallpaper.png");

      //mp1.add(mbb1, mbb2, mbb3);
      mb1.add(mbb1, mbb2, mbb3);

      this.add(mb1);

      // Font Combo
      var comboFont = new qx.ui.form.ComboBox();

      var item;


      item = new qx.ui.form.ListItem("Arial", null, "arial");
      item.setFont(new qx.ui.core.Font(16, ["arial"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Arial Black", null, "arial black");
      item.setFont(new qx.ui.core.Font(16, ["arial black"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Comic sans ms", null, "comic sans ms");
      item.setFont(new qx.ui.core.Font(16, ["comic sans ms"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Courier", null, "courier");
      item.setFont(new qx.ui.core.Font(16, ["courier"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Courier New", null, "courier new");
      item.setFont(new qx.ui.core.Font(16, ["courier new"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Cursive", null, "cursive");
      item.setFont(new qx.ui.core.Font(16, ["cursive"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Fantasy", null, "fantasy");
      item.setFont(new qx.ui.core.Font(16, ["fantasy"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Georgia", null, "georgia");
      item.setFont(new qx.ui.core.Font(16, ["georgia"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Helvetiva", null, "helvetiva");
      item.setFont(new qx.ui.core.Font(16, ["helvetiva"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Impact", null, "impact");
      item.setFont(new qx.ui.core.Font(16, ["impact"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Monospace", null, "monospace");
      item.setFont(new qx.ui.core.Font(16, ["monospace"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Palatino", null, "palatino");
      item.setFont(new qx.ui.core.Font(16, ["palatino"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Sans-serif", null, "sans-serif");
      item.setFont(new qx.ui.core.Font(16, ["sans-serif"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Serif", null, "serif");
      item.setFont(new qx.ui.core.Font(16, ["serif"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Tahoma", null, "tahoma");
      item.setFont(new qx.ui.core.Font(16, ["tahoma"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Times new roman", null, "times new roman");
      item.setFont(new qx.ui.core.Font(16, ["times new roman"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Trebuchet MS", null, "trebuchet ms");
      item.setFont(new qx.ui.core.Font(16, ["trebuchet ms"]));
      comboFont.add(item);
      item = new qx.ui.form.ListItem("Verdana", null, "verdana");
      item.setFont(new qx.ui.core.Font(16, ["verdana"]));
      comboFont.add(item);

      comboFont.setSelected(comboFont.getList().getFirstChild());

      comboFont.addEventListener("changeSelected", function(e)
      {
        this.debug("Font Combo selected: " + e.getValue());
        prevThis.changeFont(e);
      });


      // New Document Button
      var newDocBtn = new qx.ui.toolbar.Button("", "./resource_spec/icon/sp_newdoc.png");
      newDocBtn.addEventListener("execute", function(e)
      {
        this.debug("New Doc button pressed: ");
        prevThis.cdeNew.execute();
      });
      
      // Open Button
      var openBtn = new qx.ui.toolbar.Button("", "./resource_spec/icon/sp_open.png");
      openBtn.addEventListener("execute", function(e)
      {
        this.debug("Open button pressed: ");
        prevThis.cdeOpen.execute();
      });
      
      // Save Button
      var saveBtn = new qx.ui.toolbar.Button("", "./resource_spec/icon/sp_save.png");
      saveBtn.addEventListener("execute", function(e)
      {
        this.debug("Save button pressed: ");
        prevThis.cdeSave.execute();
      });
      
      // SaveAs Button
      var saveAsBtn = new qx.ui.toolbar.Button("", "./resource_spec/icon/sp_saveas.png");
      saveAsBtn.addEventListener("execute", function(e)
      {
        this.debug("Save As button pressed: ");
        prevThis.cdeSaveAs.execute();
      });
      
      // Close Button
      var closeBtn = new qx.ui.toolbar.Button("", "./resource_spec/icon/sp_close.png");
      closeBtn.addEventListener("execute", function(e)
      {
        this.debug("Close button pressed: ");
        prevThis.cdeClose.execute();
      });


      // Font size Combo
      var comboFontSize = new qx.ui.form.ComboBox();

      item = new qx.ui.form.ListItem("10", null, "10");
      comboFontSize.add(item);
      item = new qx.ui.form.ListItem("12", null, "12");
      comboFontSize.add(item);
      item = new qx.ui.form.ListItem("14", null, "14");
      comboFontSize.add(item);
      item = new qx.ui.form.ListItem("16", null, "16");
      comboFontSize.add(item);
      item = new qx.ui.form.ListItem("18", null, "18");
      comboFontSize.add(item);
      item = new qx.ui.form.ListItem("24", null, "24");
      comboFontSize.add(item);
      item = new qx.ui.form.ListItem("32", null, "32");
      comboFontSize.add(item);
      item = new qx.ui.form.ListItem("40", null, "40");
      comboFontSize.add(item);

      comboFontSize.setSelected(comboFontSize.getList().getFirstChild());

      comboFontSize.addEventListener("changeSelected", function(e)
      {
        this.debug("Font Size Combo selected: " + e.getValue());
        prevThis.changeFontSize(e);
      });

      comboFontSize.setWidth(10);

      // Align
      this.alignLeft = new qx.ui.toolbar.RadioButton("", "./resource_spec/icon/sp_alignleft.png");
      this.alignCenter = new qx.ui.toolbar.RadioButton("", "./resource_spec/icon/sp_aligncenter.png");
      this.alignRight = new qx.ui.toolbar.RadioButton("", "./resource_spec/icon/sp_alignright.png");
      this.alignJustify = new qx.ui.toolbar.RadioButton("", "./resource_spec/icon/sp_alignjustify.png");

      var alignRadMgr = new qx.ui.selection.RadioManager(null, [ this.alignLeft, this.alignCenter, this.alignRight, this.alignJustify ]);

      alignRadMgr.addEventListener("changeSelected", function(e)
      {
        this.debug("Align selected: " + e.getValue());

        var align = "";

        if (prevThis.alignLeft.getChecked()) {
          align = "left";
        } else if (prevThis.alignCenter.getChecked()) {
          align = "center";
        } else if (prevThis.alignRight.getChecked()) {
          align = "right";
        } else if (prevThis.alignJustify.getChecked()) {
          align = "justify";
        }

        prevThis.changeAlign(align);
      });

      // Font Style (Bold..)
      this.styleBold = new qx.ui.toolbar.CheckBox("", "./resource_spec/icon/sp_bold.png");

      this.styleBold.addEventListener("changeChecked", function(e)
      {
        this.debug("styleBold selected: " + e.getValue());
        prevThis.makeBold(e);
      });
      

      this.styleItalic = new qx.ui.toolbar.CheckBox("", "./resource_spec/icon/sp_italic.png");

      this.styleItalic.addEventListener("changeChecked", function(e)
      {
        this.debug("styleItalic selected: " + e.getValue());
        prevThis.makeItalic(e);
      });

      this.styleUnderline = new qx.ui.toolbar.CheckBox("", "./resource_spec/icon/sp_underline.png");

      this.styleUnderline.addEventListener("changeChecked", function(e)
      {
        this.debug("styleUnderline selected: " + e.getValue());
        prevThis.makeUnderline(e);
      });

      // Font Color chooser
      var fontColorBtn = new qx.ui.toolbar.Button("", "./resource_spec/icon/sp_color.png");

      fontColorBtn.addEventListener("execute", function(e)
      {
        this.debug("Color Font pressed: ");

        popup = new qx.ui.popup.Popup();

        var colorSel = new qx.ui.component.ColorSelector();

        colorSel.addEventListener("dialogok", function(e)
        {
          prevThis.cliDoc.remove(popup);
          this.debug("Evenement changement couleur : " + this.getRed() + "," + this.getGreen() + "," + this.getBlue() + " ");
//          fontColorBtn.setBackgroundColor(this.getRed(), this.getGreen(), this.getBlue());
          var color = "rgb(" + this.getRed() + "," + this.getGreen() + "," + this.getBlue() + ")";
          prevThis.changeColor(color);
        });

        colorSel.addEventListener("dialogcancel", function(e) {
          prevThis.cliDoc.remove(popup);
        });

        popup.setTop(qx.html.Location.getPageBoxBottom(fontColorBtn.getElement()));
        popup.setLeft(qx.html.Location.getPageBoxLeft(fontColorBtn.getElement()));
        colorSel.setBackgroundColor(fontColorBtn.getBackgroundColor());
        
        popup.add(colorSel);
        popup.addToDocument();
        popup.show();
        popup.bringToFront();

      });

      // Background Color chooser
      var backgroundColorBtn = new qx.ui.toolbar.Button("", "./resource_spec/icon/sp_backgroundcolor.png");

      backgroundColorBtn.addEventListener("execute", function(e)
      {
        this.debug("Background Color pressed: ");
        
        popup = new qx.ui.popup.Popup();

        var colorSel = new qx.ui.component.ColorSelector();

        colorSel.addEventListener("dialogok", function(e)
        {
          prevThis.cliDoc.remove(popup);
          this.debug("Evenement changement couleur : " + this.getRed() + "," + this.getGreen() + "," + this.getBlue() + " ");
//          backgroundColorBtn.setBackgroundColor(this.getRed(), this.getGreen(), this.getBlue());
          var color = "rgb(" + this.getRed() + "," + this.getGreen() + "," + this.getBlue() + ")";
          prevThis.changeBGColor(color);
        });

        colorSel.addEventListener("dialogcancel", function(e) {
            prevThis.cliDoc.remove(popup);
        });

        popup.setTop(qx.html.Location.getPageBoxBottom(backgroundColorBtn.getElement()));
        popup.setLeft(qx.html.Location.getPageBoxLeft(backgroundColorBtn.getElement()));
        colorSel.setBackgroundColor(backgroundColorBtn.getBackgroundColor());
        
        popup.add(colorSel);

        popup.addToDocument();
        popup.show();
        popup.bringToFront();


      });


      var tbMain = new qx.ui.toolbar.ToolBar;
      tbMain.setWidth("100%");
      var tbpMain1 = new qx.ui.toolbar.Part;
      tbMain.add(tbpMain1);
      tbpMain1.add(newDocBtn);
      tbpMain1.add(openBtn);
      tbpMain1.add(saveBtn);
      tbpMain1.add(saveAsBtn);
      tbpMain1.add(closeBtn);

      var tb1 = new qx.ui.toolbar.ToolBar;
      tb1.setWidth("100%");
      var tbp1 = new qx.ui.toolbar.Part;
      var tbp2 = new qx.ui.toolbar.Part;
      var tbp3 = new qx.ui.toolbar.Part;
      var tbp4 = new qx.ui.toolbar.Part;
      var tbp5 = new qx.ui.toolbar.Part;

      tb1.add(tbp1);
      tb1.add(tbp2);
      tb1.add(tbp3);
      tb1.add(tbp4);
      tb1.add(tbp5);

      tbp1.add(comboFont);
      tbp1.add(comboFontSize);

      tbp2.add(this.alignLeft);
      tbp2.add(this.alignCenter);
      tbp2.add(this.alignRight);
      tbp2.add(this.alignJustify);

      tbp3.add(this.styleBold);
      tbp3.add(this.styleItalic);
      tbp3.add(this.styleUnderline);

      tbp4.add(fontColorBtn);
      tbp4.add(backgroundColorBtn);

      this.add(tbMain);
      this.add(tb1);
    },


    /**
     * Creates the tab container which will hold the sheets
     *
     * @type member
     * @return {void} 
     */
    createTabContainer : function()
    {
      this.tabView = new qx.ui.pageview.tabview.TabView;

      this.set(
      {
        left   : 5,
        top    : 5,
        right  : 5,
        bottom : 5
      });

      this.tabView.setWidth("100%");
      this.tabView.setHeight("75%");

      this.tabView.setPlaceBarOnTop(false);

      this.add(this.tabView);
    },


    /**
     * Callback for row menu
     *
     * @type member
     * @param item {var} the menu item
     * @param val {var} the value for the item selected
     * @return {void} 
     */
    rowAction : function(item, val)
    {
      if (val != "-1")
      {
        switch(val)
        {
          case "1":  // insert row before
            this.tabs[this.activeTab].sheet.insertRowBefore();
            break;

          case "2":  // insert row before
            this.tabs[this.activeTab].sheet.insertRowAfter();
            break;

          case "3":  // insert row before
            this.tabs[this.activeTab].sheet.removeRows();
            break;
        }

        this.rowsMenu.domNode.getElementsByTagName("select")[0].selectedIndex = 0;
      }
    },


    /**
     * Callback for column menu
     *
     * @type member
     * @param item {var} the menu item
     * @param val {var} the value for the item selected
     * @return {void} 
     */
    colAction : function(item, val)
    {
      if (val != "-1")
      {
        switch(val)
        {
          case "1":  // insert row before
            this.tabs[this.activeTab].sheet.insertColumnBefore();
            break;

          case "2":  // insert row before
            this.tabs[this.activeTab].sheet.insertColumnAfter();
            break;

          case "3":  // insert row before
            this.tabs[this.activeTab].sheet.removeCols();
            break;
        }

        this.rowsMenu.domNode.getElementsByTagName("select")[0].selectedIndex = 0;
      }
    },


    /**
     * Callback for sheet menu
     *
     * @type member
     * @param item {var} the menu item
     * @param val {var} the value for the item selected
     * @return {void} 
     */
    sheetAction : function(item, val)
    {
      if (val != "-1")
      {
        switch(val)
        {
          case "1":  // rename current sheet
            var newVal = prompt("Enter the new name for the sheet");

            if (newVal)
            {
              this.tabs[this.activeTab].div.getElementsByTagName("span")[0].innerHTML = newVal;
              this.tabs[this.activeTab].label = newVal;
            }

            break;

          case "2":  // delete current sheet
            this.removeCurrentSheet();
            break;

          case "3":  // new sheet
            this.createSheet("sheet " + (this.tabs.length + 1));
            break;
        }

        this.sheetMenu.domNode.getElementsByTagName("select")[0].selectedIndex = 0;
      }
    },


    /**
     * Cut the current selection
     *
     * @type member
     * @return {void} 
     */
    cmdCut : function() {
      this.tabs[this.activeTab].sheet.cmdCutCells();
    },


    /**
     * Copy the current selection
     *
     * @type member
     * @return {void} 
     */
    cmdCopy : function() {
      this.tabs[this.activeTab].sheet.cmdCopyCells();
    },


    /**
     * Paste the current selection
     *
     * @type member
     * @return {void} 
     */
    cmdPaste : function() {
      this.tabs[this.activeTab].sheet.cmdPasteCells();
    },


    /**
     * Delete the current cells selection
     *
     * @type member
     * @return {void} 
     */
    cmdDelCells : function() {
      this.tabs[this.activeTab].sheet.cmdDeleteCells();
    },

    /**
     * Delete the current Rows selection
     *
     * @type member
     * @return {void} 
     */
    cmdDelRows : function() {
      this.tabs[this.activeTab].sheet.removeRows();
    },

    /**
     * Delete the current Columns selection
     *
     * @type member
     * @return {void} 
     */
    cmdDelColumns : function() {
      this.tabs[this.activeTab].sheet.removeCols();
    },


    /**
     * Insert Row
     *
     * @type member
     * @return {void} 
     */
    cmdInsertRow : function() {
      this.tabs[this.activeTab].sheet.insertRowBefore();
    },

    /**
     * Insert Column
     *
     * @type member
     * @return {void} 
     */
    cmdInsertColumn : function() {
      this.tabs[this.activeTab].sheet.insertColumnBefore();
    },

    /**
     * Edit formula
     *
     * @type member
     * @return {void} 
     */
    cmdEditFormula : function() {
      this.tabs[this.activeTab].sheet.cmdEditFormula();
    },


    /**
     * Show the "About" window
     *
     * @type member
     * @return {void}
     */
    cmdAbout : function() {

      var winAbout = new qx.ui.window.Window("About X4VIEW SpreadSheet");
      winAbout.setTop(200);
      winAbout.setLeft(200);
      winAbout.setWidth(200);
      winAbout.setHeight(150);

      winAbout.setShowMaximize(false);
      winAbout.setShowMinimize(false);
      winAbout.setAllowMaximize(false);
      winAbout.setAllowMinimize(false);
      winAbout.setModal(true);
      winAbout.setMoveMethod("translucent");
      winAbout.addToDocument();
      var btnOk = new qx.ui.form.Button("OK");
      btnOk.setWidth(60);
      btnOk.addEventListener("execute", function(e) {
         winAbout.fadeOut(40, 30);
         winAbout.addEventListener("FADE_FINISHED", function(e) {
           winAbout.close();
         });
      });
      var horLay = new qx.ui.layout.HorizontalBoxLayout();
      horLay.auto();
      horLay.setSpacing(10);
      horLay.set(
      {
        left   : 100,
        right  : 0
      });
      horLay.setHorizontalChildrenAlign("right");
      horLay.add(btnOk);

      var vertLay = new qx.ui.layout.VerticalBoxLayout();
      vertLay.auto();
      vertLay.set(
      {
        left   : 10,
        top    : 10,
        right  : 10,
        bottom : 10
      });
      vertLay.setSpacing(40);
      
      var atomInfo = new qx.ui.basic.Atom("", "icon/32/actions/help-about.png");
      var labelVersion = new qx.ui.basic.Label(this.version);
      var labelAuthor = new qx.ui.basic.Label(this.author);

      vertLay.add(atomInfo);
      vertLay.add(labelVersion);
      vertLay.add(labelAuthor);
      vertLay.add(horLay);
      winAbout.add(vertLay);
      winAbout.setOpacity(0);
      winAbout.fadeIn(40, 30);
      winAbout.open();




    },


    /**
     * Applies a formula
     *
     * @type member
     * @param funct {var} TODOC
     * @return {void} 
     */
    applyFunction : function(funct) {
      this.tabs[this.activeTab].sheet.applyFormula(funct);
    },


    /**
     * Callback for font size menu
     *
     * @type member
     * @param e {Event} TODOC
     * @return {void} 
     */
    changeFontSize : function(e) {
      this.tabs[this.activeTab].sheet.formatSelectedCells(this.FORMATTING_TYPES.FONT_SIZE, e.getData().getValue());
    },


    /**
     * Callback for font size menu
     *
     * @type member
     * @param e {Event} TODOC
     * @return {void} 
     */
    changeFont : function(e) {
      this.tabs[this.activeTab].sheet.formatSelectedCells(this.FORMATTING_TYPES.FONT, e.getData().getValue());
    },


    /**
     * Callback for color picker
     *
     * @type member
     * @param color {var} the color selected
     * @return {void} 
     */
    changeColor : function(color) {
      this.tabs[this.activeTab].sheet.formatSelectedCells(this.FORMATTING_TYPES.COLOR, color);
    },


    /**
     * Callback for background color picker
     *
     * @type member
     * @param color {var} the color selected
     * @return {void} 
     */
    changeBGColor : function(color) {
      this.tabs[this.activeTab].sheet.formatSelectedCells(this.FORMATTING_TYPES.BG_COLOR, color);
    },


    /**
     * Callback for bold button
     *
     * @type member
     * @param e {Event} TODOC
     * @return {void} 
     */
    makeBold : function(e) {
      this.tabs[this.activeTab].sheet.formatSelectedCells(this.FORMATTING_TYPES.BOLD, e.getValue());
    },


    /**
     * Callback for italic button
     *
     * @type member
     * @param e {Event} TODOC
     * @return {void} 
     */
    makeItalic : function(e) {
      this.tabs[this.activeTab].sheet.formatSelectedCells(this.FORMATTING_TYPES.ITALIC, e.getValue());
    },


    /**
     * Callback for underline button
     *
     * @type member
     * @param e {Event} TODOC
     * @return {void} 
     */
    makeUnderline : function(e) {
      this.tabs[this.activeTab].sheet.formatSelectedCells(this.FORMATTING_TYPES.UNDERLINE, e.getValue());
    },


    /**
     * Callback for Text Align
     *
     * @type member
     * @param align {var} the selected align
     * @return {void} 
     */
    changeAlign : function(align) {
      this.tabs[this.activeTab].sheet.formatSelectedCells(this.FORMATTING_TYPES.ALIGN, align);
    },


    /**
     * Creates/Open a new sheet (add a new sheet inside the current TabViewPage)
     *
     * @type member
     * @param sheetNbr {var} The Sheet Number
     * @param sheetName {var} The Sheet Name
     * @param sheetHtml {var} The Sheet data (XML/HTML)
     * @param additionalData {var} Additional DOM meta Data saved for the sheet
     * @return {void} 
     */
    initSheet : function(sheetNbr, sheetName, sheetHtml, additionalData)
    {
      
      // for New Sheet or Open existing Sheet
      var openExisting = false;
      if (sheetHtml != "") {
        openExisting = true;
      }

      // create the content pane for the new tab for this new sheet
      var currentTabIdx = this.tabs.length;

      var tabViewButton = new qx.ui.pageview.tabview.Button(sheetName);
      tabViewButton.setShowCloseButton(true);

      var prevThis = this;

      // add an eventlistener on the tab buttons (listen for closing tabs)
      tabViewButton.addEventListener("closetab", function(e)
      {
        this.debug("Sheet " + currentTabIdx + " Closed");
        prevThis.tabView.getBar().remove(prevThis.tabViewButtons[currentTabIdx]);
        prevThis.tabView.getPane().remove(prevThis.tabViewPages[currentTabIdx]);
        prevThis.tabViewPages[currentTabIdx].dispose();

        prevThis.createDispatchDataEvent("closesheet", ""+currentTabIdx);


      });

      tabViewButton.addEventListener("changeChecked", function(e)
      {
        if (this.getChecked()) {
          this.debug("Sheet " + currentTabIdx + " Selected");
        }
      });

      tabViewButton.setChecked(true);
      this.tabView.getBar().add(tabViewButton);
      var tabViewPage = new qx.ui.pageview.tabview.Page(tabViewButton);

      this.tabView.getPane().add(tabViewPage);

      this.tabViewButtons[currentTabIdx] = tabViewButton;
      this.tabViewPages[currentTabIdx] = tabViewPage;

      var vertLay = new qx.ui.layout.VerticalBoxLayout();

      vertLay.auto();

      vertLay.set(
      {
        left   : 1,
        top    : 1,
        right  : 1,
        bottom : 1
      });

      var horLay = new qx.ui.layout.HorizontalBoxLayout();
      horLay.auto();
      horLay.setWidth("90%");

      // Function chooser
      var prevThis = this;
      var funcButton = new qx.ui.form.Button("", "./resource_spec/icon/sp_insertformula.png");
      funcButton.setHeight(20);

      funcButton.addEventListener("execute", function(e)
      {
        this.debug("Function Button pressed: ");

        var winFunc = new qx.ui.window.Window("Choose a function", "icon/16/categories/preferences-system-network.png");
        winFunc.setTop(qx.html.Location.getPageBoxBottom(funcButton.getElement()));
        winFunc.setLeft(qx.html.Location.getPageBoxLeft(funcButton.getElement()));
        winFunc.setWidth(100);
        winFunc.setHeight(200);
        winFunc.setAllowMaximize(false);
        winFunc.setAllowMinimize(false);
        winFunc.setMoveMethod("translucent");
        prevThis.cliDoc.add(winFunc);

        // Function Combo
        var comboFunc = new qx.ui.form.ComboBox();

        var item;

        item = new qx.ui.form.ListItem("Sum", null, "sum");
        comboFunc.add(item);
        item = new qx.ui.form.ListItem("Product", null, "product");
        comboFunc.add(item);
        item = new qx.ui.form.ListItem("Average", null, "avg");
        comboFunc.add(item);
        item = new qx.ui.form.ListItem("Min", null, "min");
        comboFunc.add(item);
        item = new qx.ui.form.ListItem("Max", null, "max");
        comboFunc.add(item);
        item = new qx.ui.form.ListItem("Count", null, "count");
        comboFunc.add(item);
        item = new qx.ui.form.ListItem("Flash Movie", null, "flashmovie");
        comboFunc.add(item);
        item = new qx.ui.form.ListItem("Web page", null, "webpage");
        comboFunc.add(item);
        item = new qx.ui.form.ListItem("PDF Document", null, "pdffile");
        comboFunc.add(item);
        item = new qx.ui.form.ListItem("Word Document", null, "wordfile");
        comboFunc.add(item);
        item = new qx.ui.form.ListItem("Excel Document", null, "excelfile");
        comboFunc.add(item);


        comboFunc.setSelected(comboFunc.getList().getFirstChild());

        var btnOk = new qx.ui.form.Button("OK");
        btnOk.setWidth(60);

        btnOk.addEventListener("execute", function(e)
        {
          this.debug("Function selected: " + comboFunc.getList().getSelectedItem().getValue());
          var funct = comboFunc.getList().getSelectedItem().getValue();
          
          // Now try to propose automatically a sum (or other function) of cells above the current cell
          var sheet = prevThis.tabs[currentTabIdx].sheet;
          var curCellRow = sheet.currentFocusedRow;
          var curCellCol = sheet.currentFocusedCol;
          var startCellRow = 0;
          var endCellRow = 0;
          if (curCellRow > 0) {
             endCellRow = curCellRow - 1;
          }
          var startCellNotation = sheet.fromArrToCellNotation(curCellCol, startCellRow);
          var endCellNotation = sheet.fromArrToCellNotation(curCellCol, endCellRow);

          if (funct == "flashmovie") {
             prevThis.formulaField.setValue("=" + funct + "(\"" + "http://www.dailymotion.com/swf/1a1Hb25UtqnjK3usd" + "\")");
             //prevThis.formulaField.setValue("=" + funct + "(\"" + "http://www.youtube.com/v/o6xTNU-DWrU" + "\")");
          }
          else if (funct == "webpage") {
             prevThis.formulaField.setValue("=" + funct + "(\"" + "http://www.dailymotion.com" + "\")");
          }
          else if (funct == "pdffile") {
             prevThis.formulaField.setValue("=" + funct + "(\"" + "http://www.mywebserver.com/mypdfdocument.pdf" + "\")");
          }
          else if (funct == "wordfile") {
             prevThis.formulaField.setValue("=" + funct + "(\"" + "http://www.mywebserver.com/myworddocument.doc" + "\")");
          }
          else if (funct == "excelfile") {
             prevThis.formulaField.setValue("=" + funct + "(\"" + "http://www.mywebserver.com/myexceldocument.xls" + "\")");
          }
          else {
             prevThis.formulaField.setValue("=" + funct + "(" + startCellNotation + ":" + endCellNotation + ")");
          }



          // prevThis.applyFunction(funct);
          winFunc.close();
          prevThis.formulaField.focus();
        });

        var btnCancel = new qx.ui.form.Button("Cancel");
        btnCancel.setWidth(60);

        btnCancel.addEventListener("execute", function(e)
        {
          winFunc.close();
          prevThis.formulaField.focus();
        });

        var horLay = new qx.ui.layout.HorizontalBoxLayout();
        horLay.auto();
        horLay.setLeft(20);
        horLay.setSpacing(10);
        horLay.add(btnOk);
        horLay.add(btnCancel);

        var vertLay = new qx.ui.layout.VerticalBoxLayout();

        vertLay.auto();

        vertLay.set(
        {
          left   : 10,
          top    : 10,
          right  : 10,
          bottom : 10
        });

        vertLay.setSpacing(100);
        vertLay.add(comboFunc);
        vertLay.add(horLay);
        winFunc.add(vertLay);
        winFunc.open();
      });

      horLay.add(funcButton);

      this.formulaField = new qx.ui.form.TextField();
      this.formulaField.setWidth("100%");
      this.formulaField.setBackgroundColor("white");
      this.formulaField.setBorder("inset-thin");
      this.tabViewFormula[currentTabIdx] = this.formulaField;
      horLay.add(this.formulaField);

      vertLay.add(horLay);

      var contentHtml = new qx.ui.embed.HtmlEmbed("");
      if (openExisting == true) {
        contentHtml.setHtml(sheetHtml);
      }

      contentHtml.setWidth("100%");
      contentHtml.setHeight("90%");
      contentHtml.setOverflow("auto");

      // Context menu (right clic on spreadsheet)
      var ctxMnu = new qx.ui.menu.Menu;
      var ctx_01 = new qx.ui.menu.Button("Edit formula", "./resource_spec/icon/sp_insertformula.png", this.cdeEditFormula);
      var ctx_01S = new qx.ui.menu.Separator();
      var ctx_02 = new qx.ui.menu.Button("Cut", "icon/16/actions/edit-cut.png", this.cdeCut);
      var ctx_03 = new qx.ui.menu.Button("Copy", "icon/16/actions/edit-copy.png", this.cdeCopy);
      var ctx_04 = new qx.ui.menu.Button("Paste", "icon/16/actions/edit-paste.png", this.cdePast);

      var ctx_05S = new qx.ui.menu.Separator();

      var ctxMnuDelete = new qx.ui.menu.Menu;

      var ctx_05_01 = new qx.ui.menu.Button("Delete selected Cells", null, this.cdeDelCells);
      var ctx_05_02 = new qx.ui.menu.Button("Delete selected Rows", null, this.cdeDelRows);
      var ctx_05_03 = new qx.ui.menu.Button("Delete selected Columns", null, this.cdeDelColumns);
      ctxMnuDelete.add(ctx_05_01, ctx_05_02, ctx_05_03);
      
      var ctx_05 = new qx.ui.menu.Button("Delete", "icon/16/actions/edit-delete.png", null, ctxMnuDelete);

      var ctxMnuInsert = new qx.ui.menu.Menu;

      var ctx_06_01 = new qx.ui.menu.Button("Insert Row", null, this.cdeInsertRow);
      var ctx_06_02 = new qx.ui.menu.Button("Insert Column", null, this.cdeInsertColumn);
      ctxMnuInsert.add(ctx_06_01, ctx_06_02);

      var ctx_06 = new qx.ui.menu.Button("Insert", null, null, ctxMnuInsert);

      ctxMnu.add(ctx_01, ctx_01S, ctx_02, ctx_03, ctx_04, ctx_05, ctx_05S, ctx_06);
      
      ctxMnu.addToDocument();
      ctxMnuDelete.addToDocument();
      ctxMnuInsert.addToDocument();

      contentHtml.addEventListener("contextmenu", function(e)
      {
        ctxMnu.setLeft(e.getPageX());
        ctxMnu.setTop(e.getPageY());
        ctxMnu.show();
      });

      var sheetName = ("sheet" + currentTabIdx);
      this.tabs[currentTabIdx] = contentHtml;
      vertLay.add(contentHtml);
      this.tabViewPages[currentTabIdx].add(vertLay);

      var sheet = new qx.ui.spread.SpreadSheet(sheetName);

      contentHtml.addEventListener("insertDom", function(e) {
        sheet.initSheet(contentHtml, contentHtml.getElement(), prevThis.formulaField, funcButton, prevThis, openExisting, additionalData);
      });

      this.tabs[currentTabIdx].sheet = sheet;
      this.activeTab = currentTabIdx;
    },


    /**
     * Removes the current working sheet
     *
     * @type member
     * @return {void} 
     */
    removeCurrentSheet : function()
    {
      if (this.tabs.length > 1)
      {
        this.tabContainer.removeChild(this.tabs[this.activeTab]);

        for (var i=this.activeTab; i<this.tabs.length-1; i++) {
          this.tabs[i] = this.tabs[i + 1];
        }

        this.tabs.length--;
        this.activeTab = this.activeTab > 0 ? this.activeTab - 1 : this.activeTab;
      }
      else
      {
        alert("The document must have at least one sheet");
      }
    },


    /**
     * Handles sheet change. This is needed for the sheet to decouple certain events that may mess up how sheet receive events
     *
     * @type member
     * @return {void} 
     */
    tabChanged : function()
    {
      for (var i=0; i<this.tabs.length; i++)
      {
        if (this.tabs[i] == this.tabContainer.selectedTabWidget)
        {
          this.activeTab = i;

          if (this.tabs[i].sheet) {
            this.tabs[i].sheet.gainFocus();
          }
        }
        else
        {
          if (this.tabs[i].sheet) {
            this.tabs[i].sheet.loseFocus();
          }
        }
      }
    }
  }
});
