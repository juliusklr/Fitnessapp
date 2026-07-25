import Toybox.Graphics;
import Toybox.Lang;
import Toybox.WatchUi;

// Log sets for one exercise: adjust KG / WDH, START logs the set.
// UP/DOWN (or swipe) = ±, long UP (menu) = switch field, touch = select/adjust.
class SetView extends WatchUi.View {
    var day;
    var item;
    var exName;

    var kg = 0.0;      // 0 = bodyweight → gewicht null
    var wdh = 10;
    var satz = 1;
    var field = 1;     // 0 = KG, 1 = WDH
    var busy = false;
    var touched = false;
    var lastLine = null;
    var loaded = false;

    // box hit areas, computed in onUpdate
    var boxY = 0;
    var boxH = 0;
    var scrW = 0;

    function initialize(dayIdx, itemIdx) {
        View.initialize();
        day = gModel.days[dayIdx];
        item = day["items"][itemIdx];
        exName = item["exercise"] == null ? "?" : item["exercise"]["name"];
        var zw = firstInt(item["ziel_wdh"]);
        if (zw != null) { wdh = zw; }
    }

    function onShow() {
        if (!loaded) {
            loaded = true;
            gApi.request(:get, "/rest/v1/log_sets",
                {
                    "exercise_id" => "eq." + item["exercise_id"],
                    "select" => "datum,satz,gewicht,wdh",
                    "order" => "datum.desc,satz.desc",
                    "limit" => "12",
                },
                null, method(:onHistory));
        }
    }

    // first integer inside e.g. "8-12" / "6-12 e/s" → 8 / 6
    function firstInt(s) {
        if (s == null) { return null; }
        var str = "" + s;
        var n = null;
        var chars = str.toCharArray();
        for (var i = 0; i < chars.size(); i++) {
            var c = chars[i];
            if (c >= '0' && c <= '9') {
                n = (n == null ? 0 : n * 10) + (c.toString().toNumber());
            } else if (n != null) {
                break;
            }
        }
        return n;
    }

    function onHistory(ok, data) {
        if (!ok || !(data instanceof Lang.Array)) { return; }
        var lastDate = null;
        var sets = [];
        for (var i = 0; i < data.size(); i++) {
            var r = data[i];
            var datum = r["datum"];
            if (datum == null) { continue; }
            if (datum.equals(gModel.today)) {
                var s = r["satz"];
                if (s != null && s + 1 > satz) { satz = s + 1; }
            } else {
                if (lastDate == null) { lastDate = datum; }
                if (datum.equals(lastDate)) { sets.add(r); }
            }
        }
        if (sets.size() > 0) {
            // rows are satz-descending → last element is set 1
            var first = sets[sets.size() - 1];
            if (!touched) {
                if (first["gewicht"] != null) { kg = first["gewicht"].toFloat(); }
                if (first["wdh"] != null) { wdh = first["wdh"].toNumber(); }
            }
            var parts = "";
            for (var i = sets.size() - 1; i >= 0; i--) {
                var r = sets[i];
                var p = "";
                if (r["gewicht"] != null) { p = fmtKg(r["gewicht"].toFloat()) + "×"; }
                p += r["wdh"] == null ? "?" : r["wdh"].toNumber().toString();
                parts += (parts.equals("") ? "" : " ") + p;
            }
            lastLine = gModel.shortDate(lastDate) + " " + parts;
        }
        WatchUi.requestUpdate();
    }

    function fmtKg(f) {
        var n = f.toNumber();
        if (f == n.toFloat()) { return n.toString(); }
        return f.format("%.1f");
    }

    function adjust(dir) {
        touched = true;
        if (field == 0) {
            kg += dir * 2.5;
            if (kg < 0) { kg = 0.0; }
        } else {
            wdh += dir;
            if (wdh < 0) { wdh = 0; }
        }
        WatchUi.requestUpdate();
    }

    function toggleField() {
        field = 1 - field;
        WatchUi.requestUpdate();
    }

    function logSet() {
        if (busy) { return; }
        busy = true;
        WatchUi.requestUpdate();
        gApi.request(:post, "/rest/v1/log_sets", null,
            {
                "datum" => gModel.today,
                "plan_name" => day["label"],
                "plan_id" => day["id"],
                "exercise_id" => item["exercise_id"],
                "exercise_name" => exName,
                "satz" => satz,
                "gewicht" => kg > 0 ? kg : null,
                "wdh" => wdh,
            },
            method(:onLogged));
    }

    function onLogged(ok, data) {
        busy = false;
        if (ok) {
            WatchUi.showToast("Satz " + satz + " ✓", null);
            satz += 1;
        } else {
            WatchUi.showToast("Fehler: " + data, null);
        }
        WatchUi.requestUpdate();
    }

    function onUpdate(dc) {
        dc.setColor(Graphics.COLOR_BLACK, Graphics.COLOR_BLACK);
        dc.clear();
        var w = dc.getWidth();
        var h = dc.getHeight();
        scrW = w;

        // exercise name (truncated to fit)
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        var name = exName;
        while (dc.getTextWidthInPixels(name, Graphics.FONT_TINY) > w - 40 && name.length() > 4) {
            name = name.substring(0, name.length() - 2) + "…";
        }
        dc.drawText(w / 2, h * 0.10, Graphics.FONT_TINY, name, Graphics.TEXT_JUSTIFY_CENTER);

        // set number
        dc.setColor(0xD81413, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w / 2, h * 0.20, Graphics.FONT_XTINY, "SATZ " + satz, Graphics.TEXT_JUSTIFY_CENTER);

        // value boxes
        var bw = (w - 70) / 2;
        boxH = h * 0.30;
        boxY = h * 0.34;
        var xKg = w / 2 - bw - 8;
        var xWdh = w / 2 + 8;
        drawBox(dc, xKg, boxY, bw, boxH, "KG", kg > 0 ? fmtKg(kg) : "—", field == 0);
        drawBox(dc, xWdh, boxY, bw, boxH, "WDH", wdh.toString(), field == 1);

        // last session
        dc.setColor(Graphics.COLOR_LT_GRAY, Graphics.COLOR_TRANSPARENT);
        if (lastLine != null) {
            var ll = lastLine;
            while (dc.getTextWidthInPixels(ll, Graphics.FONT_XTINY) > w - 40 && ll.length() > 4) {
                ll = ll.substring(0, ll.length() - 2) + "…";
            }
            dc.drawText(w / 2, h * 0.70, Graphics.FONT_XTINY, ll, Graphics.TEXT_JUSTIFY_CENTER);
        }

        dc.setColor(busy ? 0xD81413 : Graphics.COLOR_DK_GRAY, Graphics.COLOR_TRANSPARENT);
        dc.drawText(w / 2, h * 0.80, Graphics.FONT_XTINY,
            busy ? "Speichern…" : "START = loggen", Graphics.TEXT_JUSTIFY_CENTER);
    }

    function drawBox(dc, x, y, bw, bh, label, value, selected) {
        dc.setPenWidth(selected ? 3 : 1);
        dc.setColor(selected ? 0xD81413 : Graphics.COLOR_DK_GRAY, Graphics.COLOR_TRANSPARENT);
        dc.drawRoundedRectangle(x, y, bw, bh, 10);
        dc.setColor(Graphics.COLOR_LT_GRAY, Graphics.COLOR_TRANSPARENT);
        dc.drawText(x + bw / 2, y + 6, Graphics.FONT_XTINY, label, Graphics.TEXT_JUSTIFY_CENTER);
        dc.setColor(Graphics.COLOR_WHITE, Graphics.COLOR_TRANSPARENT);
        dc.drawText(x + bw / 2, y + bh / 2 - 4, Graphics.FONT_NUMBER_MILD, value, Graphics.TEXT_JUSTIFY_CENTER);
    }

    // touch: tap a box to select it; tap the selected box (top/bottom half) to ±
    function handleTap(x, y) {
        var tappedField = x < scrW / 2 ? 0 : 1;
        if (y < boxY || y > boxY + boxH) { return false; }
        if (tappedField != field) {
            field = tappedField;
        } else {
            adjust(y < boxY + boxH / 2 ? 1 : -1);
        }
        WatchUi.requestUpdate();
        return true;
    }
}

class SetDelegate extends WatchUi.BehaviorDelegate {
    var view;

    function initialize(v) {
        BehaviorDelegate.initialize();
        view = v;
    }

    function onPreviousPage() { // UP / swipe down
        view.adjust(1);
        return true;
    }

    function onNextPage() { // DOWN / swipe up
        view.adjust(-1);
        return true;
    }

    function onSelect() { // START
        view.logSet();
        return true;
    }

    function onMenu() { // long UP
        view.toggleField();
        return true;
    }

    function onTap(evt) {
        var c = evt.getCoordinates();
        return view.handleTap(c[0], c[1]);
    }
}
