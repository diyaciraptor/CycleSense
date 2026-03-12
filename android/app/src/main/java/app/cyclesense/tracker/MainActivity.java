package app.cyclesense.tracker;

import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.util.Base64;
import android.webkit.JavascriptInterface;
import android.widget.Toast;

import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.core.content.FileProvider;

import com.getcapacitor.BridgeActivity;

import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

public class MainActivity extends BridgeActivity {
    private ActivityResultLauncher<String[]> importPickerLauncher;

    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        importPickerLauncher = registerForActivityResult(new ActivityResultContracts.OpenDocument(), uri -> {
            if (uri == null) {
                return;
            }

            try {
                String base64Content = readUriAsBase64(uri);
                String fileName = resolveFileName(uri);
                String script = "window.handleNativeImportXlsx(" + JSONObject.quote(base64Content) + ", " + JSONObject.quote(fileName) + ");";
                runOnUiThread(() -> bridge.getWebView().evaluateJavascript(script, null));
            } catch (Exception exception) {
                runOnUiThread(() -> Toast.makeText(MainActivity.this, "Import failed. Please try again.", Toast.LENGTH_LONG).show());
            }
        });

        bridge.getWebView().addJavascriptInterface(new ExportBridge(), "AndroidExport");
        bridge.getWebView().addJavascriptInterface(new ImportBridge(), "AndroidImport");
    }

    private class ExportBridge {
        @JavascriptInterface
        public void saveXlsx(String base64Content, String fileName) {
            try {
                byte[] bytes = Base64.decode(base64Content, Base64.DEFAULT);
                File exportDir = new File(getCacheDir(), "exports");
                if (!exportDir.exists() && !exportDir.mkdirs()) {
                    throw new IOException("Could not create export directory.");
                }

                File exportFile = new File(exportDir, sanitizeFileName(fileName));
                try (FileOutputStream outputStream = new FileOutputStream(exportFile)) {
                    outputStream.write(bytes);
                }

                runOnUiThread(() -> shareExport(exportFile));
            } catch (Exception exception) {
                runOnUiThread(() -> Toast.makeText(MainActivity.this, "Export failed. Please try again.", Toast.LENGTH_LONG).show());
            }
        }
    }

    private class ImportBridge {
        @JavascriptInterface
        public void pickXlsx() {
            runOnUiThread(() -> importPickerLauncher.launch(new String[]{
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/octet-stream"
            }));
        }
    }

    private void shareExport(File exportFile) {
        Uri uri = FileProvider.getUriForFile(this, getPackageName() + ".fileprovider", exportFile);
        Intent shareIntent = new Intent(Intent.ACTION_SEND);
        shareIntent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        shareIntent.putExtra(Intent.EXTRA_STREAM, uri);
        shareIntent.putExtra(Intent.EXTRA_SUBJECT, "CycleSense export");
        shareIntent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        startActivity(Intent.createChooser(shareIntent, "Export CycleSense data"));
    }

    private String readUriAsBase64(Uri uri) throws IOException {
        try (InputStream inputStream = getContentResolver().openInputStream(uri);
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            if (inputStream == null) {
                throw new IOException("Could not open selected file.");
            }

            byte[] buffer = new byte[8192];
            int read;
            while ((read = inputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, read);
            }

            return Base64.encodeToString(outputStream.toByteArray(), Base64.NO_WRAP);
        }
    }

    private String resolveFileName(Uri uri) {
        String lastSegment = uri.getLastPathSegment();
        if (lastSegment == null || lastSegment.isEmpty()) {
            return "cyclesense-import.xlsx";
        }
        return sanitizeFileName(lastSegment);
    }

    private String sanitizeFileName(String fileName) {
        String cleaned = fileName == null ? "cyclesense-export.xlsx" : fileName.replaceAll("[^A-Za-z0-9._-]", "_");
        if (!cleaned.endsWith(".xlsx")) {
            cleaned = cleaned + ".xlsx";
        }
        return cleaned;
    }
}