import * as $ from 'jquery';


export class UiBuilder {

  public static DisplayNotification(message: string): void {

    $("#notification").fadeIn("slow").append(message);
    $(".dismiss").click(function () {
      $("#notification").fadeOut("slow");
    });

  }

  public static AddTopNavLink(caption: string, command: Function): void {


    var link = $("<a>", { href: "javascript:void(0)" })
      .addClass("nav-link")
      .addClass("text-dark")
      .text(caption)
      .click(() => { command(); });

    var button: JQuery = $("<li>").addClass("nav-item").append(link);

    var toolbar: JQuery = $("#top-nav");
    toolbar.append(button);

  }


  public static AddCommandButton(caption: string, command: Function) : void {   
    

    var link = $("<a>", { href: "javascript:void(0)" })
      .addClass("nav-link")
      .addClass("text-dark")
      .text(caption)
      .click(() => { command(); });

    var button: JQuery = $("<li>").addClass("nav-item").append(link);

    var toolbar: JQuery = $("#embed-toolbar");
    toolbar.append(button);

  }

  public static ClearEmbedToolbar(): void {
    var toolbar: JQuery = $("#embed-toolbar");
    toolbar.empty();
  }

  private static leftNavigationEnabled: boolean = false;

  public static EnabledLeftNavigation(): void {

    var leftNavWidth: string = "180px";

    $("#embed-container").css({ "margin-left": leftNavWidth });

    var leftNav: JQuery = $("<div>", { id: "left-nav" }).css({ float: "left", width: leftNavWidth });

    var leftNavItemContainer: JQuery = $("<ul>", { id: "left-nav-item-container" }).css({ width: "100%" });

    leftNav.append(leftNavItemContainer);

    $("#content-box").prepend(leftNav);
     

    this.leftNavigationEnabled = true;

  }

  public static EnsureLeftNavigation(): void {
    if (!this.leftNavigationEnabled) {
      this.EnabledLeftNavigation();
    }
  }

  public static AddLeftNavigationHeader(caption: string): void {
    this.EnsureLeftNavigation();

    var leftNavHeader: JQuery =
      $("<li>")
        .text(caption).addClass("navbar-header");

    $("#left-nav-item-container").append(leftNavHeader);

  }

  public static AddLeftNavigationItem(caption: string, command: Function): void {
    this.EnsureLeftNavigation();

    var leftNavItem: JQuery =
      $("<li>")
        .addClass("nav-item")
        .append(
          $("<a>", { href: "javascript: void(0);" })
            .text(caption)
            .addClass("nav-link"))
            .click(() => {
              command();
            });

    $("#left-nav-item-container").append(leftNavItem);

  }
}
