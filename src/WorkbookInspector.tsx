import { useEffect, useRef, useState, type ReactNode } from "react";
import {
  indexToColumn,
  type CellDataType,
  type RangeAddress,
} from "@ricsam/formula-engine";
import { StyleActions, type FormattingProps } from "./StyleActions";
import { useActiveSelection } from "./workspace-selection";
import {
  applyDataType,
  applyStyle,
  firstAddress,
  saveTable,
  setDimension,
  validateTableName,
} from "./formatting";
import { Button } from "./components/ui/button";
import { Input } from "./components/ui/input";
import { Label } from "./components/ui/label";
import { Tabs, TabsList, TabsTrigger, TabsContent } from "./components/ui/tabs";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "./components/ui/select";
import { Switch } from "./components/ui/switch";
import { Separator } from "./components/ui/separator";

function Field({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="grid gap-1.5">
      <Label>{label}</Label>
      {children}
    </div>
  );
}
export function WorkbookInspector(props: FormattingProps) {
  const { engine, selection, active, revision, onError } = props;
  const areas = useActiveSelection(selection, active);
  const focusedTargets = useRef(
    new WeakMap<HTMLInputElement, RangeAddress[]>(),
  );
  const captureTargets = (event: React.FocusEvent<HTMLInputElement>) => {
    focusedTargets.current.set(event.currentTarget, structuredClone(areas));
  };
  const targetsFor = (input: HTMLInputElement) =>
    focusedTargets.current.get(input) ?? [];
  const address = firstAddress(areas);
  const style = address ? engine.getCellStyle(address) : undefined;
  const disabled = !active || !areas.length;
  const run = (action: () => void) => {
    try {
      if (!props.isLocked?.()) action();
    } catch (error) {
      onError(error);
    }
  };
  const [tableName, setTableName] = useState("");
  const [newName, setNewName] = useState("Table1");
  const tables = active
    ? [...engine.getTables(active.workbookName).values()].filter(
        (t) => t.sheetName === active.sheetName,
      )
    : [];
  const table = tables.find((t) => t.name === tableName);
  const selectedTable = address
    ? engine.isCellInTable(address)?.name
    : undefined;
  useEffect(() => {
    setTableName(selectedTable ?? "");
  }, [active?.id, active?.workbookName, active?.sheetName, selectedTable]);
  useEffect(() => {
    if (table) setNewName(table.name);
  }, [table?.name]);
  const metadata = active
    ? (engine.getSheetMetadata(active) as
        | {
            arenaColumnWidths?: Record<string, number>;
            arenaRowHeights?: Record<string, number>;
          }
        | undefined)
    : undefined;
  const inputKey = `${active?.id}:${address?.rowIndex}:${address?.colIndex}:${revision}`;
  const rangeLabel = table
    ? `${indexToColumn(table.start.colIndex)}${table.start.rowIndex + 1}:${indexToColumn(table.start.colIndex + table.headers.size - 1)}${table.endRow.type === "number" ? table.endRow.value + 1 : "∞"}`
    : "";
  return (
    <Tabs
      defaultValue="cell"
      className="workbook-inspector flex min-h-0 flex-1 flex-col"
    >
      <TabsList className="m-3 grid grid-cols-3">
        <TabsTrigger value="cell">Cell</TabsTrigger>
        <TabsTrigger value="table">Table</TabsTrigger>
        <TabsTrigger value="sheet">Sheet</TabsTrigger>
      </TabsList>
      <TabsContent value="cell" className="m-0 space-y-4 overflow-auto p-4">
        <Field label="Text style">
          <StyleActions {...props} />
        </Field>
        <Separator />
        <Field label="Font family">
          <Select
            disabled={disabled}
            value={style?.fontFamily ?? "DM Sans"}
            onValueChange={(fontFamily) =>
              run(() => applyStyle(engine, areas, { fontFamily }))
            }
          >
            <SelectTrigger aria-label="Font family">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              {[
                ...new Set(
                  [
                    style?.fontFamily,
                    "DM Sans",
                    "Arial",
                    "Georgia",
                    "Times New Roman",
                    "IBM Plex Mono",
                  ].filter((v): v is string => Boolean(v)),
                ),
              ].map((font) => (
                <SelectItem key={font} value={font}>
                  {font}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </Field>
        <Field label="Font size">
          <Input
            key={`font:${inputKey}`}
            aria-label="Font size"
            type="number"
            min={6}
            max={144}
            disabled={disabled}
            defaultValue={style?.fontSize ?? 11}
            onFocus={captureTargets}
            onBlur={(e) => {
              if (e.target.value && e.target.value !== e.target.defaultValue)
                run(() =>
                  applyStyle(engine, targetsFor(e.currentTarget), {
                    fontSize: Number(e.target.value),
                  }),
                );
            }}
            onKeyDown={(e) => {
              if (e.key === "Enter") e.currentTarget.blur();
            }}
          />
        </Field>
        <div className="flex items-center justify-between">
          <Label htmlFor="inspector-wrap">Wrap text</Label>
          <Switch
            id="inspector-wrap"
            disabled={disabled}
            checked={Boolean(style?.wrapText)}
            onCheckedChange={(wrapText) =>
              run(() => applyStyle(engine, areas, { wrapText }))
            }
          />
        </div>
        <Separator />
        <Field label="Data type">
          <Select
            disabled={disabled}
            value={address ? engine.getCellDataType(address) : "general"}
            onValueChange={(value) =>
              run(() => applyDataType(engine, areas, value as CellDataType))
            }
          >
            <SelectTrigger aria-label="Data type">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              {["general", "text", "number", "boolean"].map((type) => (
                <SelectItem key={type} value={type}>
                  {type[0].toUpperCase() + type.slice(1)}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </Field>
      </TabsContent>
      <TabsContent value="table" className="m-0 space-y-4 overflow-auto p-4">
        <Field label="Table">
          <Select
            disabled={!active}
            value={table ? table.name : "__new__"}
            onValueChange={(value) => {
              setTableName(value === "__new__" ? "" : value);
              if (value === "__new__") setNewName("Table1");
            }}
          >
            <SelectTrigger aria-label="Table">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="__new__">New table</SelectItem>
              {tables.map((t) => (
                <SelectItem key={t.name} value={t.name}>
                  {t.name}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </Field>
        <Field label="Table name">
          <Input
            aria-label="Table name"
            disabled={!active}
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
          />
        </Field>
        {table && (
          <Field label="Table range">
            <Input aria-label="Table range" readOnly value={rangeLabel} />
          </Field>
        )}
        {table ? (
          <div className="grid gap-2">
            <Button
              variant="outline"
              disabled={!active || !newName.trim()}
              onClick={() =>
                run(() => {
                  const name = newName.trim();
                  validateTableName(
                    engine,
                    table.workbookName,
                    name,
                    table.name,
                  );
                  if (name !== table.name)
                    engine.renameTable(table.workbookName, {
                      oldName: table.name,
                      newName: name,
                    });
                  setTableName(name);
                })
              }
            >
              Rename table
            </Button>
            <Button
              variant="outline"
              disabled={disabled}
              onClick={() =>
                run(() => saveTable(engine, areas, table.name, table))
              }
            >
              Resize to selection
            </Button>
            <Button
              variant="outline"
              onClick={() =>
                run(() => {
                  engine.removeTable({
                    workbookName: table.workbookName,
                    tableName: table.name,
                  });
                  setTableName("");
                })
              }
            >
              Remove table definition
            </Button>
          </div>
        ) : (
          <Button
            className="w-full"
            disabled={disabled || !newName.trim()}
            onClick={() =>
              run(() => {
                saveTable(engine, areas, newName);
                setTableName(newName.trim());
              })
            }
          >
            Create from selection
          </Button>
        )}
        {table && (
          <div className="grid grid-cols-2 gap-2">
            <Field label="Columns">
              <Input
                aria-label="Table columns"
                readOnly
                value={table.headers.size}
              />
            </Field>
            <Field label="Data rows">
              <Input
                aria-label="Table data rows"
                readOnly
                value={
                  table.endRow.type === "number"
                    ? table.endRow.value - table.start.rowIndex
                    : "∞"
                }
              />
            </Field>
          </div>
        )}
      </TabsContent>
      <TabsContent value="sheet" className="m-0 space-y-4 overflow-auto p-4">
        <Field label="Column width">
          <Input
            key={`col:${inputKey}`}
            aria-label="Column width"
            onFocus={captureTargets}
            type="number"
            min={16}
            max={1000}
            disabled={disabled}
            defaultValue={
              address
                ? (metadata?.arenaColumnWidths?.[
                    indexToColumn(address.colIndex)
                  ] ?? 100)
                : 100
            }
            onBlur={(e) => {
              if (e.target.value && e.target.value !== e.target.defaultValue)
                run(() =>
                  setDimension(
                    engine,
                    targetsFor(e.currentTarget),
                    "col",
                    Number(e.target.value),
                  ),
                );
            }}
            onKeyDown={(e) => {
              if (e.key === "Enter") e.currentTarget.blur();
            }}
          />
        </Field>
        <Field label="Row height">
          <Input
            key={`row:${inputKey}`}
            aria-label="Row height"
            onFocus={captureTargets}
            type="number"
            min={16}
            max={500}
            disabled={disabled}
            defaultValue={
              address
                ? (metadata?.arenaRowHeights?.[String(address.rowIndex + 1)] ??
                  28)
                : 28
            }
            onBlur={(e) => {
              if (e.target.value && e.target.value !== e.target.defaultValue)
                run(() =>
                  setDimension(
                    engine,
                    targetsFor(e.currentTarget),
                    "row",
                    Number(e.target.value),
                  ),
                );
            }}
            onKeyDown={(e) => {
              if (e.key === "Enter") e.currentTarget.blur();
            }}
          />
        </Field>
      </TabsContent>
    </Tabs>
  );
}
