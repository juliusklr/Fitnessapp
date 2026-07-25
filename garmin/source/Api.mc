import Toybox.Application;
import Toybox.Communications;
import Toybox.Lang;
import Toybox.System;

// Supabase REST layer. Auth via GoTrue password grant, data via PostgREST.
// All requests go through request(); on a 401 the token is refreshed once
// and the original request is replayed.
class Api {
    var url;
    var key;
    var token = null;

    // pending request (for the 401 replay)
    var _method = null;
    var _path = null;
    var _params = null;
    var _body = null;
    var _cb = null;
    var _retried = false;
    var _loginCb = null;

    function initialize() {
        url = Application.Properties.getValue("SupabaseUrl");
        key = Application.Properties.getValue("SupabaseKey");
    }

    function hasCredentials() {
        var e = Application.Properties.getValue("Email");
        return e != null && !e.equals("") && !e.equals("__SUPABASE_EMAIL__");
    }

    // ── login ──
    function login(cb) {
        _loginCb = cb;
        Communications.makeWebRequest(
            url + "/auth/v1/token?grant_type=password",
            {
                "email" => Application.Properties.getValue("Email"),
                "password" => Application.Properties.getValue("Password"),
            },
            {
                :method => Communications.HTTP_REQUEST_METHOD_POST,
                :headers => {
                    "Content-Type" => Communications.REQUEST_CONTENT_TYPE_JSON,
                    "apikey" => key,
                },
                :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON,
            },
            method(:onLogin)
        );
    }

    function onLogin(code, data) {
        var cb = _loginCb;
        _loginCb = null;
        if (code == 200 && data != null && data["access_token"] != null) {
            token = data["access_token"];
            if (cb != null) { cb.invoke(true, null); }
        } else {
            token = null;
            if (cb != null) { cb.invoke(false, errText(code, data)); }
        }
    }

    // ── generic PostgREST request ──
    // method: :get | :post   path: "/rest/v1/..."   params: query dict or null
    // body: dict or null   cb.invoke(ok, dataOrError)
    function request(httpMethod, path, params, body, cb) {
        _method = httpMethod;
        _path = path;
        _params = params;
        _body = body;
        _cb = cb;
        _retried = false;
        _send();
    }

    function _send() {
        var headers = {
            "apikey" => key,
            "Authorization" => "Bearer " + token,
        };
        var opts = {
            :headers => headers,
            :responseType => Communications.HTTP_RESPONSE_CONTENT_TYPE_JSON,
        };
        var payload = _params;
        if (_method == :post) {
            opts[:method] = Communications.HTTP_REQUEST_METHOD_POST;
            headers["Content-Type"] = Communications.REQUEST_CONTENT_TYPE_JSON;
            headers["Prefer"] = "return=representation";
            payload = _body;
        } else {
            opts[:method] = Communications.HTTP_REQUEST_METHOD_GET;
        }
        var full = url + _path;
        if (_method == :post && _params != null) {
            // query params for POST must live in the URL (payload is the body)
            var qs = "";
            var keys = _params.keys();
            for (var i = 0; i < keys.size(); i++) {
                qs += (i == 0 ? "?" : "&") + keys[i] + "=" + Communications.encodeURL(_params[keys[i]]);
            }
            full += qs;
        }
        Communications.makeWebRequest(full, payload, opts, method(:onResponse));
    }

    function onResponse(code, data) {
        if (code == 401 && !_retried) {
            _retried = true;
            login(method(:onRelogin));
            return;
        }
        var cb = _cb;
        _cb = null;
        if (cb == null) { return; }
        if (code >= 200 && code < 300) {
            cb.invoke(true, data);
        } else {
            cb.invoke(false, errText(code, data));
        }
    }

    function onRelogin(ok, err) {
        if (ok) {
            _send();
        } else {
            var cb = _cb;
            _cb = null;
            if (cb != null) { cb.invoke(false, err); }
        }
    }

    function errText(code, data) {
        if (code == -104) { return "Kein Handy?"; }
        if (data instanceof Lang.Dictionary) {
            var m = data["message"];
            if (m == null) { m = data["error_description"]; }
            if (m == null) { m = data["msg"]; }
            if (m != null) { return m + " (" + code + ")"; }
        }
        return "Fehler " + code;
    }
}
