// ── ObjectType mappings ──────────────────────────────────────────────────────

/** Map from our friendly type names to VML ObjectType values */
export const VML_OBJ_TYPE: Record<string, string> = {
  button:       'Button',
  checkBox:     'Checkbox',
  comboBox:     'Drop',
  listBox:      'List',
  optionButton: 'Radio',
  groupBox:     'GBox',
  label:        'Label',
  scrollBar:    'Scroll',
  spinner:      'Spin',
  dialog:       'Dialog',
};

/** Map from our friendly type names to ctrlProp objectType values (differs from VML for CheckBox) */
export const CTRL_OBJ_TYPE: Record<string, string> = {
  ...VML_OBJ_TYPE,
  checkBox:     'CheckBox',   // ctrlProp uses capital B
  dialog:       'Dialog',
};

/** Reverse map: OOXML objectType → our type name (handles both VML and ctrlProp casing) */
export const OBJ_TYPE_TO_CTRL: Record<string, string> = {
  ...Object.fromEntries(Object.entries(VML_OBJ_TYPE).map(([k, v]) => [v, k])),
  'CheckBox': 'checkBox',  // ctrlProp uses capital B
};

/** Checked state to OOXML numeric */
export const CHECKED_MAP: Record<string, string> = {
  unchecked: '0', checked: '1', mixed: '2',
};
export const CHECKED_REV: Record<string, string> = { '0': 'unchecked', '1': 'checked', '2': 'mixed' };
