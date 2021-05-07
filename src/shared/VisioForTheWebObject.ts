export class VisioForTheWebObject {

  private visioUrlDocument = "";
  private visioEmbeddedSession: OfficeExtension.EmbeddedSession = null;
  private fullFileUniqueId: string;

  // ========================================================
  // Get the sharePoint file Unique Id of the Visio file including the path
  //  docUrl url of the visio document (Copied from url address of the browser when viewing a Visio Document)
  //  return the fullFileuniqueId
  // ========================================================
  public getFullFileUniqueId(url: string): string {
    var index = url.indexOf("&file=");
    var fullFileuniqueId = url.substr(0, index);
    return fullFileuniqueId;
  }

  // ========================================================
  //  initializes the Visio embed session and thus display the diagram
  //  in the iFrame
  //  docUrl url of the visio document (Copied from url address of the browser when viewing a Visio Document)
  //  returns a promise
  // ========================================================
  public load = async (docUrl: string): Promise<void> => {
    try {
      // get the fileUniqueId
      this.fullFileUniqueId = this.getFullFileUniqueId(docUrl);
      // build the session Url for being able to call Visio JS API function
      this.visioUrlDocument = this.fullFileUniqueId + "&action=embedview";
      // call initialization of Visio embedded session
      await this.CreateVisioEmbeddedSession();
      this.addCustomEventHandlers();
    } catch (error) {
      this.logError(error);
    }
  }

  // ========================================================
  // initialize an office embedded session with the iFrame of the SharePoint page
  // returns a promise
  // ========================================================
  private async CreateVisioEmbeddedSession(): Promise<any> {
    // Remove previous created element for preventing to have multiple iFrame
    let docRootElement = document.getElementById("iframeHost");
    if (docRootElement != null) {
      for (var i = 0; i < docRootElement.childNodes.length; ++i) {
        docRootElement.removeChild(docRootElement.childNodes[i]);
      }
      // the created visioEmbeddedSession will be used for all further call to Visio for the web component
      this.visioEmbeddedSession = new OfficeExtension.EmbeddedSession(
        this.visioUrlDocument, {
        id: "visioembedded-iframe",
        container: document.getElementById("iframeHost"),
        width: "100%",
        height: "600px"
      }
      );
      await this.visioEmbeddedSession.init();
    }
  }

  // ========================================================
  // PlaceHolder for adding API function
  // ========================================================
  public VisioJSApiFunction = async (): Promise<void> => {
    try {
      await Visio.run(this.visioEmbeddedSession, async (context: Visio.RequestContext) => {
        // Add your Visio JS Api calls
        await context.sync();
      });
    } catch (error) {
      this.logError(error);
    }
  }

  // ========================================================
  // Highlight a Shape
  // shapeName name of the shape to higlight
  // bHighlight higlight if true un-highlight if false
  // returns a promise
  // ========================================================
  public highlightShape = async (shapeName: string, bHighlight: boolean): Promise<void> => {
    try {
      await Visio.run(this.visioEmbeddedSession, async (context: Visio.RequestContext) => {
        const activePage: Visio.Page = context.document.getActivePage();
        const shapesCollection: Visio.ShapeCollection = activePage.shapes;
        const shape: Visio.Shape = shapesCollection.getItem(shapeName);
        if (bHighlight == true)
          shape.view.highlight = { color: "#FF0000", width: 2 };
        else
          shape.view.highlight = null;
        await context.sync();
      });
    } catch (error) {
      this.logError(error);
    }
  }



  // ========================================================
  // Add custom event handlers
  // returns a promise
  // ========================================================
  private addCustomEventHandlers = async (): Promise<any> => {

    try {
      await Visio.run(this.visioEmbeddedSession, async (context: Visio.RequestContext) => {
        var visioDocument: Visio.Document = context.document;
        // on mouse enter
        const onShapeMouseEnterEventResult: OfficeExtension.EventHandlerResult<Visio.ShapeMouseEnterEventArgs> =
          visioDocument.onShapeMouseEnter.add(this.onShapeMouseEnter);
        await context.sync();
      });
    } catch (error) {
      this.logError(error);
    }
  }

  // ========================================================
  // PlaceHolder for subscribing to Visio Js API Event
  // ========================================================  
  private VisioJSApiEvent = async (args: Visio.ShapeMouseEnterEventArgs): Promise<void> => {
    try {
      console.log("Event fired");
    }
    catch (error) {
      this.logError(error);
    }
  }

  // ========================================================
  // ShapeMouseEnter event
  // args
  // returns a promise
  // ========================================================
  private onShapeMouseEnter = async (args: Visio.ShapeMouseEnterEventArgs): Promise<void> => {
    try {
      console.log("onShapeMouseEnter");
    }
    catch (error) {
      this.logError(error);
    }
  }

  // ========================================================
  // log error to debug console
  // error error object
  // ========================================================
  private logError = (error: any): void => {
    console.error("Error");
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info: ", JSON.stringify(error.debugInfo));
    } else {
      console.error(error);
    }
  }

}
