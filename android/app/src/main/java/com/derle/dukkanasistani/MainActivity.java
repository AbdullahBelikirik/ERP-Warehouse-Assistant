package com.derle.dukkanasistani;

import android.Manifest;
import android.app.Activity;
import android.app.AlertDialog;
import android.content.SharedPreferences;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Bundle;
import android.view.View;
import android.view.WindowManager;
import android.webkit.PermissionRequest;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceRequest;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.TextView;
import android.widget.Toast;

public class MainActivity extends Activity {
    private static final int CAMERA_PERMISSION = 10;
    private static final String PREFS = "site_settings";
    private static final String URL_KEY = "server_url";
    private WebView webView;
    private PermissionRequest cameraRequest;
    private SharedPreferences preferences;

    @Override
    public void onCreate(Bundle state) {
        super.onCreate(state);
        getWindow().addFlags(WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON);
        preferences = getSharedPreferences(PREFS, MODE_PRIVATE);
        String savedUrl = preferences.getString(URL_KEY, "");
        if (savedUrl.isEmpty()) {
            showAddressDialog("");
        } else {
            openSite(savedUrl);
        }
    }

    private void showAddressDialog(String current) {
        LinearLayout layout = new LinearLayout(this);
        layout.setOrientation(LinearLayout.VERTICAL);
        int padding = (int) (20 * getResources().getDisplayMetrics().density);
        layout.setPadding(padding, padding, padding, 0);
        TextView help = new TextView(this);
        help.setText("Ngrok HTTPS adresinizi girin. Örnek: https://ornek.ngrok-free.app");
        EditText input = new EditText(this);
        input.setSingleLine(true);
        input.setInputType(17);
        input.setHint("https://ornek.ngrok-free.app");
        input.setText(current);
        layout.addView(help);
        layout.addView(input);

        AlertDialog dialog = new AlertDialog.Builder(this)
                .setTitle("Sunucu adresi")
                .setView(layout)
                .setCancelable(false)
                .setPositiveButton("Bağlan", null)
                .create();
        dialog.setOnShowListener(ignored -> dialog.getButton(AlertDialog.BUTTON_POSITIVE).setOnClickListener(v -> {
            String url = input.getText().toString().trim();
            if (!url.startsWith("https://") && !url.startsWith("http://")) {
                url = "https://" + url;
            }
            Uri uri = Uri.parse(url);
            if (uri.getHost() == null) {
                input.setError("Geçerli bir ngrok adresi girin");
                return;
            }
            preferences.edit().putString(URL_KEY, url).apply();
            dialog.dismiss();
            openSite(url);
        }));
        dialog.show();
    }

    private void openSite(String url) {
        webView = new WebView(this);
        webView.getSettings().setJavaScriptEnabled(true);
        webView.getSettings().setDomStorageEnabled(true);
        webView.getSettings().setMediaPlaybackRequiresUserGesture(false);
        webView.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest request) {
                String host = request.getUrl().getHost();
                String siteHost = Uri.parse(preferences.getString(URL_KEY, "")).getHost();
                return host == null || !host.equals(siteHost);
            }
        });
        webView.setWebChromeClient(new WebChromeClient() {
            @Override
            public void onPermissionRequest(PermissionRequest request) {
                boolean asksForCamera = false;
                for (String resource : request.getResources()) {
                    if (PermissionRequest.RESOURCE_VIDEO_CAPTURE.equals(resource)) asksForCamera = true;
                }
                if (!asksForCamera) {
                    request.deny();
                    return;
                }
                if (checkSelfPermission(Manifest.permission.CAMERA) == PackageManager.PERMISSION_GRANTED) {
                    request.grant(new String[]{PermissionRequest.RESOURCE_VIDEO_CAPTURE});
                } else {
                    cameraRequest = request;
                    requestPermissions(new String[]{Manifest.permission.CAMERA}, CAMERA_PERMISSION);
                }
            }
        });
        setContentView(webView);
        webView.loadUrl(url);
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, String[] permissions, int[] results) {
        super.onRequestPermissionsResult(requestCode, permissions, results);
        if (requestCode == CAMERA_PERMISSION && cameraRequest != null) {
            if (results.length > 0 && results[0] == PackageManager.PERMISSION_GRANTED) {
                cameraRequest.grant(new String[]{PermissionRequest.RESOURCE_VIDEO_CAPTURE});
            } else {
                cameraRequest.deny();
                Toast.makeText(this, "Barkod taramak için kamera izni gereklidir.", Toast.LENGTH_LONG).show();
            }
            cameraRequest = null;
        }
    }

    @Override
    public void onBackPressed() {
        if (webView != null && webView.canGoBack()) {
            webView.goBack();
        } else {
            super.onBackPressed();
        }
    }

    @Override
    protected void onResume() {
        super.onResume();
        if (webView != null) webView.onResume();
    }

    @Override
    protected void onPause() {
        if (webView != null) webView.onPause();
        super.onPause();
    }
}
