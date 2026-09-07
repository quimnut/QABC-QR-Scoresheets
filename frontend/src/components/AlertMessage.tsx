import { useEffect, useRef } from "react";
import type { Alert } from "../lib/types";

interface Props {
  alerts: Alert[];
  onDismiss: (id: string) => void;
}

export default function AlertMessage({ alerts, onDismiss }: Props) {
  return (
    <>
      {alerts.map((alert) => (
        <AutoDismissAlert key={alert.id} alert={alert} onDismiss={onDismiss} />
      ))}
    </>
  );
}

function AutoDismissAlert({
  alert,
  onDismiss,
}: {
  alert: Alert;
  onDismiss: (id: string) => void;
}) {
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const timer = setTimeout(() => {
      el.style.opacity = "0";
      el.style.transition = "opacity 0.3s";
      setTimeout(() => onDismiss(alert.id), 300);
    }, 5000);
    return () => clearTimeout(timer);
  }, [alert.id, onDismiss]);

  return (
    <div ref={ref} className={`alert alert-${alert.kind}`} style={{ transition: "opacity 0.3s" }}>
      {alert.message}
    </div>
  );
}
