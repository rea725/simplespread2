qx.Class.define("ss.Application",
{
  extend : qx.application.Gui,

  members :
  {
    main : function()
    {
      // Call super class
      this.base(arguments);


      qx.Class.include(qx.ui.core.Widget, qx.ui.animation.MAnimation)
  
      qx.theme.manager.Meta.getInstance().setTheme(qx.theme.Ext);
  
      //qx.theme.manager.Icon.getInstance().setIconTheme(qx.theme.icon.VistaInspirate);
      qx.theme.manager.Icon.getInstance().setIconTheme(qx.theme.icon.Nuvola);
  
  
    	
      var d = qx.ui.core.ClientDocument.getInstance();
  
      var w0 = new qx.ui.layout.CanvasLayout;
      w0.set({left:5, top:5, right:5, bottom:5, border:"inset"});
      w0.setOverflow("hidden");
      d.add(w0); 
  
      var w1 = new qx.ui.window.Window("Simple Spreadsheet Component");
      w1.setSpace(20, 400, 20, 270);
      w0.add(w1);
  
      var tf1 = new qx.ui.pageview.tabview.TabView;
      tf1.set({ left: 10, top: 10, right: 10, bottom: 10 });
  
      var t1_1 = new qx.ui.pageview.tabview.Button("Spreadsheet demo");
  
      t1_1.setChecked(true);
  
      tf1.getBar().add(t1_1);
  
      var p1_1 = new qx.ui.pageview.tabview.Page(t1_1);
  
      tf1.getPane().add(p1_1);
  
      w1.add(tf1);
  
  
  
      var spreadData = "";
  
      var spread = new ss.Spread();
      spread.createSheet("Sheet 1");
      spread.addEventListener("open", function(e) {
         this.debug("********* Document was opened");
         spread.openSpread(spreadData, "Sheet");
      });
      spread.addEventListener("save", function(e) {
         spreadData = e.getData();
         this.debug("********* Document was saved with data : " + e.getData());
      });
      spread.addEventListener("closesheet", function(e) {
         this.debug("********* Sheet " + e.getData() + "was closed");
      });
  
      p1_1.add(spread);
  		
  
      w1.open();

  }
 }
});

