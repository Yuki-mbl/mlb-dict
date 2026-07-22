package jp.fantamstick.mlbdict;

import android.os.Bundle;
import com.getcapacitor.BridgeActivity;

public class MainActivity extends BridgeActivity {
    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        // 端末の「フォントサイズ（文字拡大）」設定に影響されず、常に標準サイズで表示する
        this.bridge.getWebView().getSettings().setTextZoom(100);
    }
}
