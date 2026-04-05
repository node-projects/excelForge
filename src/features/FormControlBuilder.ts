/**
 * FormControlBuilder — generates ctrlProp XML and VML shapes for form controls.
 */

import type { FormControl, FormControlAnchor } from '../core/types.js';
import { escapeXml } from '../utils/helpers.js';
import { CHECKED_MAP, CTRL_OBJ_TYPE, VML_OBJ_TYPE } from './FormControlBuilderCommon.js';



// ── ctrlProp XML ─────────────────────────────────────────────────────────────

export function buildCtrlPropXml(ctrl: FormControl): string {
  if (ctrl._ctrlPropXml) return ctrl._ctrlPropXml;

  const objType = CTRL_OBJ_TYPE[ctrl.type] ?? 'Button';
  const attrs: string[] = [
    `objectType="${objType}"`,
    'lockText="1"',
  ];

  if (ctrl.linkedCell) attrs.push(`fmlaLink="${escapeXml(ctrl.linkedCell)}"`);
  if (ctrl.inputRange) attrs.push(`fmlaRange="${escapeXml(ctrl.inputRange)}"`);

  // CheckBox / OptionButton
  if (ctrl.checked !== undefined) {
    attrs.push(`checked="${CHECKED_MAP[ctrl.checked] ?? '0'}"`);
  }

  // ComboBox
  if (ctrl.dropLines !== undefined) attrs.push(`dropLines="${ctrl.dropLines}"`);

  // ScrollBar / Spinner
  if (ctrl.min !== undefined) attrs.push(`min="${ctrl.min}"`);
  if (ctrl.max !== undefined) attrs.push(`max="${ctrl.max}"`);
  if (ctrl.inc !== undefined) attrs.push(`inc="${ctrl.inc}"`);
  if (ctrl.page !== undefined) attrs.push(`page="${ctrl.page}"`);
  if (ctrl.val !== undefined) attrs.push(`val="${ctrl.val}"`);

  // ListBox
  if (ctrl.selType) {
    const selMap: Record<string, string> = { single: 'Single', multi: 'Multi', extend: 'Extend' };
    attrs.push(`selType="${selMap[ctrl.selType] ?? 'Single'}"`);
  }

  if (ctrl.noThreeD) attrs.push('noThreeD="1"');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ${attrs.join(' ')}/>`;
}

// ── VML shape XML ────────────────────────────────────────────────────────────

function vmlAnchor(from: FormControlAnchor, to: FormControlAnchor): string {
  // VML anchor: fromCol, fromColOff, fromRow, fromRowOff, toCol, toColOff, toRow, toRowOff
  // Offsets are in ~15ths of a pixel for VML; we use 0 by default
  return `${from.col}, ${from.colOff ?? 0}, ${from.row}, ${from.rowOff ?? 0}, ${to.col}, ${to.colOff ?? 0}, ${to.row}, ${to.rowOff ?? 0}`;
}

export function buildFormControlVmlShape(ctrl: FormControl, shapeId: number): string {
  if (ctrl._vmlShapeXml) return ctrl._vmlShapeXml;

  const objType = VML_OBJ_TYPE[ctrl.type] ?? 'Button';
  const anchor = vmlAnchor(ctrl.from, ctrl.to);

  // Build x:ClientData children
  const cd: string[] = [];
  cd.push(`<x:Anchor>${anchor}</x:Anchor>`);
  cd.push('<x:PrintObject>False</x:PrintObject>');
  cd.push('<x:AutoFill>False</x:AutoFill>');
  if (ctrl.macro) cd.push(`<x:FmlaMacro>${escapeXml(ctrl.macro)}</x:FmlaMacro>`);
  if (ctrl.linkedCell) cd.push(`<x:FmlaLink>${escapeXml(ctrl.linkedCell)}</x:FmlaLink>`);
  if (ctrl.inputRange) cd.push(`<x:FmlaRange>${escapeXml(ctrl.inputRange)}</x:FmlaRange>`);
  if (ctrl.checked !== undefined) cd.push(`<x:Checked>${CHECKED_MAP[ctrl.checked] ?? '0'}</x:Checked>`);
  if (ctrl.dropLines !== undefined) cd.push(`<x:DropLines>${ctrl.dropLines}</x:DropLines>`);
  if (ctrl.dropStyle) cd.push(`<x:DropStyle>${escapeXml(ctrl.dropStyle)}</x:DropStyle>`);
  if (ctrl.min !== undefined) cd.push(`<x:Min>${ctrl.min}</x:Min>`);
  if (ctrl.max !== undefined) cd.push(`<x:Max>${ctrl.max}</x:Max>`);
  if (ctrl.inc !== undefined) cd.push(`<x:Inc>${ctrl.inc}</x:Inc>`);
  if (ctrl.page !== undefined) cd.push(`<x:Page>${ctrl.page}</x:Page>`);
  if (ctrl.val !== undefined) cd.push(`<x:Val>${ctrl.val}</x:Val>`);
  if (ctrl.selType) {
    const selMap: Record<string, string> = { single: 'Single', multi: 'Multi', extend: 'Extend' };
    cd.push(`<x:Sel>${selMap[ctrl.selType] ?? 'Single'}</x:Sel>`);
  }
  if (ctrl.noThreeD) cd.push('<x:NoThreeD/>');
  // Dialog-specific button attributes
  if (ctrl.isDefault) cd.push('<x:Default/>');
  if (ctrl.isDismiss) cd.push('<x:Dismiss/>');
  if (ctrl.isCancel) cd.push('<x:Cancel/>');

  // Text content for controls that display text
  let textBox = '';
  if (ctrl.text && (ctrl.type === 'button' || ctrl.type === 'checkBox' ||
      ctrl.type === 'optionButton' || ctrl.type === 'groupBox' || ctrl.type === 'label' ||
      ctrl.type === 'dialog')) {
    const align = ctrl.type === 'button' ? 'center' : 'left';
    textBox = `<v:textbox style="mso-direction-alt:auto"><div style="text-align:${align}"><font face="Calibri" size="220" color="#000000">${escapeXml(ctrl.text)}</font></div></v:textbox>`;
    cd.push('<x:TextHAlign>Center</x:TextHAlign>');
    cd.push('<x:TextVAlign>Center</x:TextVAlign>');
  }

  // Style — estimate pixel position from cell indices (approx 64px/col, 20px/row)
  const left = (ctrl.from.col * 64 + (ctrl.from.colOff ?? 0)) * 0.75;
  const top = (ctrl.from.row * 20 + (ctrl.from.rowOff ?? 0)) * 0.75;
  const width = ((ctrl.to.col - ctrl.from.col) * 64) * 0.75;
  const height = ((ctrl.to.row - ctrl.from.row) * 20) * 0.75;

  const shapeName = ctrl.text ?? `${ctrl.type}_${shapeId}`;

  // Dialog frames use a visible filled shape; other controls are invisible
  const fillStroke = ctrl.type === 'dialog' ? '' : ' filled="f" stroked="f"';

  return `<v:shape o:spid="_x0000_s${shapeId}" id="${escapeXml(shapeName)}" type="#_x0000_t201" style="position:absolute;margin-left:${left.toFixed(1)}pt;margin-top:${top.toFixed(1)}pt;width:${width.toFixed(1)}pt;height:${height.toFixed(1)}pt;z-index:${shapeId}"${fillStroke} o:insetmode="auto"><o:lock v:ext="edit" rotation="t"/>${textBox}<x:ClientData ObjectType="${objType}">${cd.join('')}</x:ClientData></v:shape>`;
}

// ── Complete VML document ────────────────────────────────────────────────────

export function buildVmlWithControls(
  commentShapes: string[],
  controlShapes: string[],
): string {
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>
<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype>
${commentShapes.join('\n')}
${controlShapes.join('\n')}
</xml>`;
}
