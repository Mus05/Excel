import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Check, ChevronDown, X } from 'lucide-react';
import { FUNCTION_CATALOG } from '../engine/formulaEngine';
import { NamedRange } from '../types/spreadsheet';

interface FormulaBarProps {
  visible: boolean;
  cellAddress: string;
  formulaText: string;
  isEditing: boolean;
  namedRanges: NamedRange[];
  onJumpToReference: (refOrName: string) => void;
  onStartEdit: () => void;
  onChangeFormula: (val: string) => void;
  onCommitEdit: () => void;
  onCancelEdit: () => void;
  onOpenFunctionWizard: () => void;
  inputRef: React.RefObject<HTMLInputElement | null>;
}

export const FormulaBar: React.FC<FormulaBarProps> = ({
  visible,
  cellAddress,
  formulaText,
  isEditing,
  namedRanges,
  onJumpToReference,
  onStartEdit,
  onChangeFormula,
  onCommitEdit,
  onCancelEdit,
  onOpenFunctionWizard,
  inputRef,
}) => {
  const [nameBoxDraft, setNameBoxDraft] = useState(cellAddress);
  const [showNameDropdown, setShowNameDropdown] = useState(false);
  const [acIndex, setAcIndex] = useState(0);
  const dropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setNameBoxDraft(cellAddress);
  }, [cellAddress]);

  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
        setShowNameDropdown(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  // Detect active function typing for autocomplete
  const activeToken = useMemo(() => {
    if (!isEditing || !formulaText.startsWith('=')) return '';
    const m = formulaText.match(/([A-Za-z_][A-Za-z0-9_]*)$/);
    return m ? m[1].toUpperCase() : '';
  }, [isEditing, formulaText]);

  const suggestions = useMemo(() => {
    if (!activeToken || activeToken.length < 1) return [];
    // Don't show autocomplete if user just finished typing a cell ref like A1 or B12
    if (/^[A-Z]{1,2}\d+$/.test(activeToken)) return [];
    return FUNCTION_CATALOG.filter((f) => f.name.startsWith(activeToken)).slice(0, 8);
  }, [activeToken]);

  // Active function syntax hint when inside parentheses e.g. =SUM(
  const activeFunctionHint = useMemo(() => {
    if (!isEditing || !formulaText.startsWith('=')) return null;
    const matches = [...formulaText.matchAll(/([A-Za-z_][A-Za-z0-9_]*)\(/g)];
    if (matches.length === 0) return null;
    const lastName = matches[matches.length - 1][1].toUpperCase();
    return FUNCTION_CATALOG.find((f) => f.name === lastName) || null;
  }, [isEditing, formulaText]);

  useEffect(() => {
    setAcIndex(0);
  }, [suggestions.length, activeToken]);

  if (!visible) return null;

  const insertSuggestion = (fnName: string) => {
    const updated = formulaText.replace(/([A-Za-z_][A-Za-z0-9_]*)$/, `${fnName}(`);
    onChangeFormula(updated);
    inputRef.current?.focus();
  };

  return (
    <div className="no-print h-8 bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-700 flex items-center px-2 gap-1.5 text-xs select-none shrink-0 relative z-20">
      {/* Name Box */}
      <div className="relative flex items-center w-28 shrink-0" ref={dropdownRef}>
        <input
          type="text"
          value={nameBoxDraft}
          onChange={(e) => setNameBoxDraft(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              onJumpToReference(nameBoxDraft);
              (e.target as HTMLInputElement).blur();
            } else if (e.key === 'Escape') {
              setNameBoxDraft(cellAddress);
              (e.target as HTMLInputElement).blur();
            }
          }}
          title="Name Box — type cell address (e.g. D10), range (A1:C5), or Named Range"
          className="w-full h-6 px-2 pr-6 font-mono-num text-xs font-medium bg-slate-50 dark:bg-slate-800 border border-slate-300 dark:border-slate-600 rounded-sm text-slate-800 dark:text-slate-100 focus:outline-none focus:border-[#107c41]"
        />
        <button
          type="button"
          onClick={() => setShowNameDropdown((v) => !v)}
          className="absolute right-0.5 top-0.5 bottom-0.5 px-1 text-slate-500 hover:text-slate-800 dark:hover:text-slate-200"
          title="Named Ranges"
        >
          <ChevronDown className="w-3 h-3" />
        </button>

        {showNameDropdown && (
          <div className="absolute left-0 top-7 w-52 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-lg py-1 z-50">
            <div className="px-2.5 py-1 text-[11px] font-semibold text-slate-400 border-b border-slate-100 dark:border-slate-700">
              Named Ranges
            </div>
            {namedRanges.length === 0 ? (
              <div className="px-2.5 py-2 text-xs text-slate-500">No named ranges defined</div>
            ) : (
              namedRanges.map((nr) => (
                <button
                  key={nr.name}
                  onClick={() => {
                    onJumpToReference(nr.name);
                    setShowNameDropdown(false);
                  }}
                  className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 dark:hover:bg-slate-700 flex items-center justify-between text-xs"
                >
                  <span className="font-medium text-slate-800 dark:text-slate-100">{nr.name}</span>
                  <span className="font-mono-num text-[11px] text-slate-500">{nr.range}</span>
                </button>
              ))
            )}
          </div>
        )}
      </div>

      {/* Cancel / Commit / fx buttons */}
      <div className="flex items-center border-x border-slate-200 dark:border-slate-700 px-1 gap-0.5 shrink-0">
        <button
          type="button"
          onClick={onCancelEdit}
          disabled={!isEditing}
          title="Cancel (Esc)"
          className="p-1 rounded hover:bg-red-50 dark:hover:bg-red-900/30 text-red-600 disabled:opacity-30 disabled:pointer-events-none"
        >
          <X className="w-3.5 h-3.5" />
        </button>
        <button
          type="button"
          onClick={onCommitEdit}
          disabled={!isEditing}
          title="Enter (Confirm)"
          className="p-1 rounded hover:bg-emerald-50 dark:hover:bg-emerald-900/30 text-[#107c41] disabled:opacity-30 disabled:pointer-events-none"
        >
          <Check className="w-3.5 h-3.5" />
        </button>
        <button
          type="button"
          onClick={onOpenFunctionWizard}
          title="Insert Function (fx)"
          className="px-1.5 py-0.5 rounded hover:bg-slate-100 dark:hover:bg-slate-800 font-serif italic font-bold text-slate-700 dark:text-slate-200 text-sm leading-none"
        >
          fx
        </button>
      </div>

      {/* Formula Input */}
      <div className="relative flex-1 min-w-0">
        <input
          ref={inputRef}
          type="text"
          value={formulaText}
          onFocus={() => {
            if (!isEditing) onStartEdit();
          }}
          onChange={(e) => onChangeFormula(e.target.value)}
          onKeyDown={(e) => {
            if (suggestions.length > 0) {
              if (e.key === 'ArrowDown') {
                e.preventDefault();
                setAcIndex((i) => (i + 1) % suggestions.length);
                return;
              }
              if (e.key === 'ArrowUp') {
                e.preventDefault();
                setAcIndex((i) => (i - 1 + suggestions.length) % suggestions.length);
                return;
              }
              if (e.key === 'Tab') {
                e.preventDefault();
                insertSuggestion(suggestions[acIndex].name);
                return;
              }
            }
            if (e.key === 'Enter') {
              e.preventDefault();
              onCommitEdit();
            } else if (e.key === 'Escape') {
              e.preventDefault();
              onCancelEdit();
            }
          }}
          placeholder="Enter value or formula starting with ="
          className="w-full h-6 px-2 font-mono-num text-xs bg-transparent text-slate-900 dark:text-slate-100 focus:outline-none"
        />

        {/* Autocomplete Dropdown */}
        {suggestions.length > 0 && (
          <div className="absolute left-0 top-7 w-80 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded shadow-lg py-1 z-50">
            {suggestions.map((fn, idx) => (
              <button
                key={fn.name}
                type="button"
                onMouseDown={(e) => {
                  e.preventDefault();
                  insertSuggestion(fn.name);
                }}
                className={`w-full px-3 py-1.5 text-left flex flex-col gap-0.5 ${
                  idx === acIndex
                    ? 'bg-emerald-50 dark:bg-emerald-900/40'
                    : 'hover:bg-slate-50 dark:hover:bg-slate-700/60'
                }`}
              >
                <div className="flex items-center justify-between">
                  <span className="font-mono-num font-semibold text-xs text-[#107c41] dark:text-emerald-400">
                    {fn.name}
                  </span>
                  <span className="text-[10px] text-slate-400">{fn.category}</span>
                </div>
                <span className="text-[11px] text-slate-500 dark:text-slate-300 truncate">
                  {fn.description}
                </span>
              </button>
            ))}
          </div>
        )}

        {/* Active Function Syntax Tooltip */}
        {activeFunctionHint && suggestions.length === 0 && (
          <div className="absolute left-0 top-7 px-2.5 py-1 bg-amber-50 dark:bg-slate-800 border border-amber-200 dark:border-slate-600 rounded shadow-sm text-[11px] font-mono-num text-slate-700 dark:text-slate-200 z-40 pointer-events-none">
            <span className="font-semibold text-[#107c41] dark:text-emerald-400">{activeFunctionHint.syntax}</span>
            <span className="ml-2 font-sans text-slate-500 dark:text-slate-400">
              — {activeFunctionHint.description}
            </span>
          </div>
        )}
      </div>
    </div>
  );
};
