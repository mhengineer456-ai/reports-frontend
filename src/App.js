// src/App.jsx
import React, { useState } from "react";
import { BrowserRouter as Router, Switch, Route, Redirect } from "react-router-dom";
import SplashScreen from "./SplashScreen";
import LoginScreen from "./LoginScreen";
import ProtectedRoute from "./ProtectedRoute";
import Dashboard from "./Dashboard";
import EmbroideryChallan from "./EmbroideryChallan";
import DailyEmbroideryChallan from "./DailyEmbroiderChallan";
import DailyPrintingChallan from "./DailyprintingChallan";
import DailyStitchingIssue from "./DailyStitchingIssue";
import PrintingChallan from "./PrintingChallan";
import CuttingStatsReport from "./Cuttingreport";
import PendingIssuetoStitching from "./PendingStitchingissue";
import PendingIssue from "./PendingIssue";
import IssueToPacking from "./IssuePacking";
import DailyPackingReport from "./DailyPackingReport";
import DailyKajButtonReport from "./DailyKajButton";
import DailyFoldingReport from "./DailyFoldingReport";
import DailyOverlockReport from "./DailyOverlock";
import DailyStitchingUpdation from "./DailyStitchingReport";
import DailyNotUpdation from "./DailyNotUpdation";
import StitchingCompleteLot from "./StitchingCompleted";
import OverallCuttingToPacking from "./OverallCuttingtoPacking";
import ShortSummarReport from "./ShortSummaryReport";
import JobOrders from "./AllJobOrder";
import ZipPurchaseDashboard from "./ZipComponent";
import DoriPurchaseDashboard from "./DoriComponent";
import StickerReport from "./StickerReport";
import PendingPackingtoIssue from "./PackingPendingtoIssue";
import PackingAlloted from "./PackingAlloted";
import EmbroideryMaterialReceiving from './EmbroideryMaterialReceiving';
import PrintingMaterialReceiving from './PrintingMaterialReceiving';
import FabricRollPrediction from './FabricRollPrediction';
import LotLogs from './LotLogs';
import NotificationCenter from './NotificationCenter';
import KnittingReport from './KnittingReport';
import CollarReport from './CollarReport';
import KoraRollReport from './KoraRollReport';
import YarnReport from './YarnReport';
import PendingZipPOReport from './PendingZipPOReport';
import EmbPrintRemarks from './EmbPrintRemarks';
import DailyFabricIssueReport from './DailyFabricIssueReport';

function App() {
  const [showSplash, setShowSplash] = useState(true);

  React.useEffect(() => {
    document.title = "MH Factory Suite Pro — Enterprise Production System";

    // Programmatically set and refresh favicon link tags to bypass aggressive browser icon caching
    const head = document.getElementsByTagName('head')[0];
    let iconLink = document.querySelector("link[rel*='icon']");
    if (!iconLink) {
      iconLink = document.createElement('link');
      iconLink.rel = 'icon';
      head.appendChild(iconLink);
    }
    iconLink.type = 'image/svg+xml';
    iconLink.href = '/favicon.svg?v=' + Date.now();
  }, []);

  // Show splash screen first
  if (showSplash) {
    return <SplashScreen onFinish={() => setShowSplash(false)} />;
  }

  return (
    <Router>
      <Switch>
        <Route exact path="/" component={LoginScreen} />
        <ProtectedRoute path="/dashboard" component={Dashboard} />
        <ProtectedRoute path="/embroidery" component={EmbroideryChallan} />
        <ProtectedRoute path="/daily-embroidery-challan" component={DailyEmbroideryChallan} />
        <ProtectedRoute path="/daily-printing-challan" component={DailyPrintingChallan} />
        <ProtectedRoute path="/stitching" component={DailyStitchingIssue} />
        <ProtectedRoute path="/printing" component={PrintingChallan} />
        <ProtectedRoute path="/cutting-report" component={CuttingStatsReport} />
        <ProtectedRoute path="/pending-stitching" component={PendingIssuetoStitching} />
        <ProtectedRoute path="/pending-issue-to-stitching" component={PendingIssue} />
        <ProtectedRoute path="/issue-to-packing" component={IssueToPacking} />
        <ProtectedRoute path="/daily-packing-report" component={DailyPackingReport} />
        <ProtectedRoute path="/daily-kaj-button-report" component={DailyKajButtonReport} />
        <ProtectedRoute path="/daily-folding-report" component={DailyFoldingReport} />
        <ProtectedRoute path="/daily-overlock-report" component={DailyOverlockReport} />
        <ProtectedRoute path="/daily-stitching-report" component={DailyStitchingUpdation} />
        <ProtectedRoute path="/daily-stitching-report-not-updation" component={DailyNotUpdation} />
        <ProtectedRoute path="/stitching-complete-lot" component={StitchingCompleteLot} />
        <ProtectedRoute path="/overall-cutting-to-packing-report" component={OverallCuttingToPacking} />
        <ProtectedRoute path="/short-summary-report" component={ShortSummarReport} adminOnly={true} />
        <ProtectedRoute path="/all-cutting-joborders" component={JobOrders} />
        <ProtectedRoute path="/zip-report" component={ZipPurchaseDashboard} />
        <ProtectedRoute path="/dori-report" component={DoriPurchaseDashboard} />
        <ProtectedRoute path="/pending-packing-issue" component={PendingPackingtoIssue} />
        <ProtectedRoute path="/sticker-report" component={StickerReport} />
        <ProtectedRoute path="/packing-alloted-lot" component={PackingAlloted} />
        <ProtectedRoute path="/embroidery-material-receiving" component={EmbroideryMaterialReceiving} />
        <ProtectedRoute path="/printing-material-receiving" component={PrintingMaterialReceiving} />
        <ProtectedRoute path="/fabric-roll-prediction" component={FabricRollPrediction} />
        <ProtectedRoute path="/lot-logs" component={LotLogs} />
        <ProtectedRoute path="/notifications" component={NotificationCenter} />
        <ProtectedRoute path="/knitting-report" component={KnittingReport} adminOnly={true} />
        <ProtectedRoute path="/collar-report" component={CollarReport} adminOnly={true} />
        <ProtectedRoute path="/kora-roll-report" component={KoraRollReport} adminOnly={true} />
        <ProtectedRoute path="/yarn-report" component={YarnReport} adminOnly={true} />
        <ProtectedRoute path="/pending-zip-po-report" component={PendingZipPOReport} />
        <ProtectedRoute path="/emb-print-remarks" component={EmbPrintRemarks} />
        <ProtectedRoute path="/daily-fabric-issue-report" component={DailyFabricIssueReport} />

        <Route path="*">
          <Redirect to="/" />
        </Route>
      </Switch>
    </Router>
  );
}

export default App;
