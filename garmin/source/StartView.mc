import Toybox.Graphics;
import Toybox.Lang;
import Toybox.WatchUi;

// Boot sequence: login → schedule(today) → weekly pattern → programs → day menu.
class StartView extends WatchUi.View {
    var status = "Verbinde…";
    var isError = false;
    var _started = false;

    function initialize() {
        View.initialize();
    }

    function onShow() {
        if (!_started) {
            _started = true;
            start();
        }
    }

    function start() {
        isError = false;
        if (!gApi.hasCredentials()) {
            showError("Keine Zugangsdaten im Build (Secrets fehlen)");
            return;
        }
        setStatus("Anmelden…");
        gApi.login(method(:onLogin));
    }

    function setStatus(s) {
        status = s;
        WatchUi.requestUpdate();
    }

    function showError(s) {
        isError = true;
        status = s;
        WatchUi.requestUpdate();
    }

    function onLogin(ok, err) {
        if (!ok) {
            showError(err);
            return;
        }
        setStatus("Lade Plan…");
        gApi.request(:get, "/rest/v1/schedule_entries",
            { "datum" => "eq." + gModel.today, "select" => "plan_id" },
            null, method(:onSchedule));
    }

    var _scheduledPlanId = null;
    var _scheduleDecided = false;

    function onSchedule(ok, data) {
        if (!ok) {
            showError(data);
            return;
        }
        if (data instanceof Lang.Array && data.size() > 0) {
            // explicit entry wins; plan_id null = deliberately free
            _scheduledPlanId = data[0]["plan_id"];
            _scheduleDecided = true;
            loadPrograms();
        } else {
            gApi.request(:get, "/rest/v1/weekly_pattern",
                { "weekday" => "eq." + gModel.weekday, "select" => "plan_id" },
                null, method(:onPattern));
        }
    }

    function onPattern(ok, data) {
        if (!ok) {
            showError(data);
            return;
        }
        if (data instanceof Lang.Array && data.size() > 0) {
            _scheduledPlanId = data[0]["plan_id"];
        }
        _scheduleDecided = true;
        loadPrograms();
    }

    function loadPrograms() {
        setStatus("Lade Übungen…");
        gApi.request(:get, "/rest/v1/programs",
            {
                "select" => "id,name,position,phases(id,name,position,plans(id,name,position,plan_items(id,exercise_id,position,gruppe,runden,ziel_saetze,ziel_wdh,exercise:exercises(name))))",
                "order" => "position",
            },
            null, method(:onPrograms));
    }

    function onPrograms(ok, data) {
        if (!ok) {
            showError(data);
            return;
        }
        if (!(data instanceof Lang.Array)) {
            showError("Unerwartete Antwort");
            return;
        }
        gModel.setPrograms(data);
        gModel.markToday(_scheduledPlanId);
        if (gModel.days.size() == 0) {
            showError("Keine Trainingstage angelegt");
            return;
        }
        pushDayMenu();
    }

    function onUpdate(dc) {
        dc.setColor(Graphics.COLOR_BLACK, Graphics.COLOR_BLACK);
        dc.clear();
        var w = dc.getWidth();
        var h = dc.getHeight();
        dc.setColor(0xD81413, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w / 2, h / 2 - 50, Graphics.FONT_MEDIUM, "Training", Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(isError ? 0xFF5550 : Graphics.COLOR_LT_GRAY, Graphics.COLOR_TRANSPARENT);
        drawWrapped(dc, status, w / 2, h / 2, w - 50);
        if (isError) {
            dc.setColor(Graphics.COLOR_DK_GRAY, Graphics.COLOR_TRANSPARENT);
            dc.drawText(w / 2, h - 55, Graphics.FONT_XTINY, "START: nochmal", Graphics.TEXT_JUSTIFY_CENTER);
        }
    }

    // naive word wrap for error messages
    function drawWrapped(dc, text, cx, y, maxW) {
        var words = splitWords(text);
        var line = "";
        var yy = y;
        for (var i = 0; i < words.size(); i++) {
            var probe = line.equals("") ? words[i] : line + " " + words[i];
            if (dc.getTextWidthInPixels(probe, Graphics.FONT_XTINY) > maxW && !line.equals("")) {
                dc.drawText(cx, yy, Graphics.FONT_XTINY, line, Graphics.TEXT_JUSTIFY_CENTER);
                yy += dc.getFontHeight(Graphics.FONT_XTINY);
                line = words[i];
            } else {
                line = probe;
            }
        }
        if (!line.equals("")) {
            dc.drawText(cx, yy, Graphics.FONT_XTINY, line, Graphics.TEXT_JUSTIFY_CENTER);
        }
    }

    function splitWords(text) {
        var out = [];
        var cur = "";
        var chars = text.toCharArray();
        for (var i = 0; i < chars.size(); i++) {
            if (chars[i] == ' ') {
                if (!cur.equals("")) { out.add(cur); }
                cur = "";
            } else {
                cur += chars[i];
            }
        }
        if (!cur.equals("")) { out.add(cur); }
        return out;
    }
}

class StartDelegate extends WatchUi.BehaviorDelegate {
    var view;

    function initialize(v) {
        BehaviorDelegate.initialize();
        view = v;
    }

    function onSelect() {
        if (view.isError) {
            view.start();
            return true;
        }
        return false;
    }
}
