import Toybox.Application;
import Toybox.Lang;
import Toybox.WatchUi;

var gApi;
var gModel;

class TrainingApp extends Application.AppBase {
    function initialize() {
        AppBase.initialize();
        gApi = new Api();
        gModel = new Model();
    }

    function onStart(state) {}
    function onStop(state) {}

    function getInitialView() {
        var v = new StartView();
        return [v, new StartDelegate(v)];
    }
}
