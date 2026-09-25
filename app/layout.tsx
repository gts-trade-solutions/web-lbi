import "./globals.css";
import AuthGate from "../components/AuthGate";
import ToastHost from "../components/Toast";
import ConfirmHost from "../components/ConfirmDialog";
import VersionWatcher from "../components/VersionWatcher";

export const metadata = {
  title: "Tracker Web",
  description: "Projects dashboard",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>
        <AuthGate>{children}</AuthGate>
        <ToastHost />
        <ConfirmHost />
        <VersionWatcher />
      </body>
    </html>
  );
}
