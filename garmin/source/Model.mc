import Toybox.Lang;
import Toybox.Time;
import Toybox.Time.Gregorian;

// Flattened training days + today's schedule resolution.
class Model {
    var days = [];        // [{ "id", "name", "label", "sub", "items" }]
    var todayIdx = null;  // index into days, if scheduled for today
    var today;            // "YYYY-MM-DD" (local)
    var weekday;          // 0 = Montag … 6 = Sonntag

    function initialize() {
        var now = Gregorian.info(Time.now(), Time.FORMAT_SHORT);
        today = now.year.format("%04d") + "-" + now.month.format("%02d") + "-" + now.day.format("%02d");
        weekday = (now.day_of_week + 5) % 7; // CIQ: 1 = Sunday … 7 = Saturday
    }

    function pos(d) {
        var p = d["position"];
        return p == null ? 0 : p;
    }

    function sortByPosition(arr) {
        for (var i = 1; i < arr.size(); i++) {
            var v = arr[i];
            var j = i - 1;
            while (j >= 0 && pos(arr[j]) > pos(v)) {
                arr[j + 1] = arr[j];
                j--;
            }
            arr[j + 1] = v;
        }
        return arr;
    }

    // programs → phases → plans (Tage) → plan_items, flattened in display order
    function setPrograms(programs) {
        days = [];
        sortByPosition(programs);
        for (var p = 0; p < programs.size(); p++) {
            var pr = programs[p];
            var phases = pr["phases"] == null ? [] : pr["phases"];
            sortByPosition(phases);
            for (var f = 0; f < phases.size(); f++) {
                var ph = phases[f];
                var plans = ph["plans"] == null ? [] : ph["plans"];
                sortByPosition(plans);
                for (var t = 0; t < plans.size(); t++) {
                    var day = plans[t];
                    var items = day["plan_items"] == null ? [] : day["plan_items"];
                    sortByPosition(items);
                    days.add({
                        "id" => day["id"],
                        "name" => day["name"],
                        "label" => pr["name"] + " · " + ph["name"] + " · " + day["name"],
                        "sub" => pr["name"] + " · " + ph["name"],
                        "items" => items,
                    });
                }
            }
        }
    }

    function markToday(planId) {
        todayIdx = null;
        if (planId == null) { return; }
        for (var i = 0; i < days.size(); i++) {
            if (planId.equals(days[i]["id"])) {
                todayIdx = i;
                return;
            }
        }
    }

    // "2026-07-21" → "21.07."
    function shortDate(iso) {
        if (iso == null || iso.length() < 10) { return ""; }
        return iso.substring(8, 10) + "." + iso.substring(5, 7) + ".";
    }
}
