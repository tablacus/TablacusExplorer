var s = ['<label>Number of items</label><br /><input type="text" id="Save" style="width: 100%" /><br /><br />'];
s.push('<label>Advanced</label><br /><input type="hidden" name="Format" />');
s.push('<input type="radio" id="Format=0" name="_Format" onclick="SetRadio(this)" checked />');
s.push('<label for="Format=0">Normal</label>');
s.push('<input type="radio" id="Format=1" name="_Format" onclick="SetRadio(this)" />');
s.push('<label for="Format=1">Path</label>');
s.push('<input type="radio" id="Format=2" name="_Format" onclick="SetRadio(this)" />');
s.push('<label for="Format=2">Key</label>');
SetTabContents(0, "General", s);
