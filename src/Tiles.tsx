import { useRef, type ReactNode } from "react";
import type { Layout, Pane } from "./model";

export function Tiles({
  layout,
  renderPane,
  onResize,
}: {
  layout: Layout;
  renderPane: (pane: Pane) => ReactNode;
  onResize: (id: string, ratio: number) => void;
}) {
  const element = useRef<HTMLDivElement>(null);
  if (layout.kind === "pane")
    return <div className="tile-leaf">{renderPane(layout.pane)}</div>;
  const horizontal = layout.direction === "horizontal";
  return (
    <div
      ref={element}
      className={`tile-split ${horizontal ? "horizontal" : "vertical"}`}
    >
      <div className="tile-branch" style={{ flex: `${layout.ratio} 1 0` }}>
        <Tiles
          layout={layout.first}
          renderPane={renderPane}
          onResize={onResize}
        />
      </div>
      <div
        className="tile-divider"
        role="separator"
        tabIndex={0}
        aria-label={`Resize ${horizontal ? "columns" : "rows"}`}
        aria-orientation={horizontal ? "vertical" : "horizontal"}
        aria-valuemin={15}
        aria-valuemax={85}
        aria-valuenow={Math.round(layout.ratio * 100)}
        onDoubleClick={() => onResize(layout.id, 0.5)}
        onKeyDown={(event) => {
          const delta =
            event.key === "ArrowLeft" || event.key === "ArrowUp"
              ? -0.05
              : event.key === "ArrowRight" || event.key === "ArrowDown"
                ? 0.05
                : 0;
          if (delta || event.key === "Home") {
            event.preventDefault();
            onResize(
              layout.id,
              event.key === "Home" ? 0.5 : layout.ratio + delta,
            );
          }
        }}
        onPointerDown={(event) => {
          event.currentTarget.setPointerCapture(event.pointerId);
          event.preventDefault();
        }}
        onPointerMove={(event) => {
          if (!event.currentTarget.hasPointerCapture(event.pointerId)) return;
          const rect = element.current!.getBoundingClientRect();
          onResize(
            layout.id,
            horizontal
              ? (event.clientX - rect.left) / rect.width
              : (event.clientY - rect.top) / rect.height,
          );
        }}
        onPointerUp={(event) =>
          event.currentTarget.releasePointerCapture(event.pointerId)
        }
      >
        <span />
      </div>
      <div className="tile-branch" style={{ flex: `${1 - layout.ratio} 1 0` }}>
        <Tiles
          layout={layout.second}
          renderPane={renderPane}
          onResize={onResize}
        />
      </div>
    </div>
  );
}
