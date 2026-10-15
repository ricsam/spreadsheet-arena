import { useId, useRef, type CSSProperties, type ReactNode } from "react";
import { cn } from "./lib/utils";

export interface ResizableSidebarProps {
  side: "left" | "right";
  width: number;
  onWidthChange: (width: number) => void;
  children: ReactNode;
  className?: string;
  "aria-label"?: string;
  minWidth?: number;
  maxWidth?: number;
}
export function ResizableSidebar({
  side,
  width,
  onWidthChange,
  children,
  className,
  "aria-label": label,
  minWidth = 220,
  maxWidth = 420,
}: ResizableSidebarProps) {
  const id = useId();
  const drag = useRef<{ x: number; width: number } | null>(null);
  const minimum = Math.max(0, minWidth);
  const maximum = Math.max(minimum, maxWidth);
  const clamp = (value: number) =>
    Math.min(
      maximum,
      Math.max(minimum, Number.isFinite(value) ? value : minimum),
    );
  const effectiveWidth = clamp(width);
  const direction = side === "left" ? 1 : -1;
  return (
    <aside
      className={cn("resizable-sidebar", className)}
      data-side={side}
      aria-label={label}
      style={{ "--sidebar-width": `${effectiveWidth}px` } as CSSProperties}
    >
      <div id={id} className="resizable-sidebar-content">
        {children}
      </div>
      <div
        className="sidebar-resize-handle"
        role="separator"
        tabIndex={0}
        aria-orientation="vertical"
        aria-label={`Resize ${label ?? `${side} sidebar`}`}
        aria-controls={id}
        aria-valuemin={minimum}
        aria-valuemax={maximum}
        aria-valuenow={effectiveWidth}
        onKeyDown={(event) => {
          let next: number;
          if (event.key === "Home") next = minimum;
          else if (event.key === "End") next = maximum;
          else if (event.key === "ArrowLeft" || event.key === "ArrowRight")
            next =
              effectiveWidth +
              (event.key === "ArrowRight" ? 1 : -1) *
                direction *
                (event.shiftKey ? 32 : 8);
          else return;
          event.preventDefault();
          onWidthChange(clamp(next));
        }}
        onPointerDown={(event) => {
          if (event.button !== 0) return;
          event.preventDefault();
          event.currentTarget.focus();
          event.currentTarget.setPointerCapture(event.pointerId);
          drag.current = { x: event.clientX, width: effectiveWidth };
        }}
        onPointerMove={(event) => {
          if (drag.current)
            onWidthChange(
              clamp(
                drag.current.width +
                  direction * (event.clientX - drag.current.x),
              ),
            );
        }}
        onPointerUp={(event) => {
          drag.current = null;
          if (event.currentTarget.hasPointerCapture(event.pointerId))
            event.currentTarget.releasePointerCapture(event.pointerId);
        }}
        onPointerCancel={() => {
          drag.current = null;
        }}
        onLostPointerCapture={() => {
          drag.current = null;
        }}
      />
    </aside>
  );
}
