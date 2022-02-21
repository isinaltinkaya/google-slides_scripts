/**
 * Credit: https://developers.google.com/apps-script/add-ons/editors/slides/quickstart/progress-bar
 */

var BAR_ID = 'PROGRESS_BAR_ID';
var BAR_TEXT = 'PROGRESS_BTEXT';

var BAR_HEIGHT = 15; // px

function onInstall(e) {
  onOpen();
}

function onOpen(e) {
  SlidesApp.getUi().createAddonMenu()
      .addItem('Show progress bar', 'createBars')
      .addItem('Hide progress bar', 'deleteBars')
      .addToUi();
}

function createBars() {
  deleteBars(); // Delete any existing progress bars
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; ++i) {
    var ratioComplete = (i / (slides.length - 1));
    var x = 0;
//    var y = presentation.getPageHeight() - BAR_HEIGHT;
//    var y = BAR_HEIGHT;
    var y=0;
    var barWidth = presentation.getPageWidth() * ratioComplete;

    if (barWidth > 0) {
      var bar = slides[i].insertShape(SlidesApp.ShapeType.RECTANGLE, x, y,
                                      barWidth, BAR_HEIGHT);
      bar.getBorder().setTransparent();
      bar.getFill().setSolidFill('#eeeeee',0.60);

      bar.setLinkUrl(BAR_ID);
      bar.sendToBack();
    }
    
    var shape = slides[i].insertShape(SlidesApp.ShapeType.TEXT_BOX, 0, -5, 100, 10);
    shape.setLinkUrl(BAR_ID);

    var textRange = shape.getText();
    var slidei= i+1;
    textRange.setText('['+slidei + '/' + slides.length+']');
    textRange.getTextStyle().setBold(false).setFontFamily("Bebas Neue").setFontSize(9).setForegroundColor('#990000');

  }

}

function deleteBars() {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; ++i) {
    var elements = slides[i].getPageElements();
    for (var j = 0; j < elements.length; ++j) {
      var el = elements[j];
      if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE &&
          el.asShape().getLink() &&
          el.asShape().getLink().getUrl() === BAR_ID) {
        el.remove();
      }
    }
  }
}
