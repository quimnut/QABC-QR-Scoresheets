interface Props {
  current: number;
  total: number;
  message: string;
  onStop?: () => void;
}

export default function ProgressBar({ current, total, message, onStop }: Props) {
  const percent = total > 0 ? Math.round((current / total) * 100) : 0;

  return (
    <div className="progress-container">
      <div className="progress-bar">
        <div className="progress-fill" style={{ width: `${percent}%` }}>
          {percent > 5 ? `${percent}%` : ""}
        </div>
      </div>
      <p className="progress-message">{message}</p>
      {onStop && (
        <button className="btn btn-danger btn-small stop-btn" onClick={onStop}>
          Stop
        </button>
      )}
    </div>
  );
}
