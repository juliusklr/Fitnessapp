import Toybox.Lang;
import Toybox.WatchUi;

// ── Day picker ──
function pushDayMenu() {
    var menu = new WatchUi.Menu2({ :title => "Training" });
    if (gModel.todayIdx != null) {
        var d = gModel.days[gModel.todayIdx];
        menu.addItem(new WatchUi.MenuItem("▶ Heute: " + d["name"], d["sub"], gModel.todayIdx, null));
    }
    for (var i = 0; i < gModel.days.size(); i++) {
        var day = gModel.days[i];
        menu.addItem(new WatchUi.MenuItem(day["name"], day["sub"], i, null));
    }
    WatchUi.switchToView(menu, new DayMenuDelegate(), WatchUi.SLIDE_IMMEDIATE);
}

class DayMenuDelegate extends WatchUi.Menu2InputDelegate {
    function initialize() {
        Menu2InputDelegate.initialize();
    }

    function onSelect(item) {
        pushExerciseMenu(item.getId());
    }
}

// ── Exercises of one day ──
function pushExerciseMenu(dayIdx) {
    var day = gModel.days[dayIdx];
    var items = day["items"];
    var menu = new WatchUi.Menu2({ :title => day["name"] });
    for (var i = 0; i < items.size(); i++) {
        var it = items[i];
        var name = it["exercise"] == null ? "?" : it["exercise"]["name"];
        var sub = "";
        if (it["gruppe"] != null && !it["gruppe"].equals("")) { sub += it["gruppe"]; }
        if (it["runden"] != null) { sub += (sub.equals("") ? "" : " · ") + it["runden"] + " Rd"; }
        if (it["ziel_saetze"] != null && !("" + it["ziel_saetze"]).equals("")) {
            sub += (sub.equals("") ? "" : " · ") + it["ziel_saetze"] + " Sätze";
        }
        if (it["ziel_wdh"] != null && !it["ziel_wdh"].equals("")) {
            sub += (sub.equals("") ? "" : " · ") + it["ziel_wdh"] + " Wdh";
        }
        menu.addItem(new WatchUi.MenuItem(name, sub.equals("") ? null : sub, i, null));
    }
    WatchUi.pushView(menu, new ExerciseMenuDelegate(dayIdx), WatchUi.SLIDE_LEFT);
}

class ExerciseMenuDelegate extends WatchUi.Menu2InputDelegate {
    var dayIdx;

    function initialize(d) {
        Menu2InputDelegate.initialize();
        dayIdx = d;
    }

    function onSelect(item) {
        var v = new SetView(dayIdx, item.getId());
        WatchUi.pushView(v, new SetDelegate(v), WatchUi.SLIDE_LEFT);
    }
}
