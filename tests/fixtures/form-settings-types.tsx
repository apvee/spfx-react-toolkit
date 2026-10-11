import * as React from 'react';
import {createForm} from '../../packages/spfx-react-toolkit/src/controls/form';
import {SettingsGroup} from '../../packages/spfx-react-toolkit/src/controls/settings-group/settings-group';
const F=createForm<{title:string;tags:string[]}>();
export const example=<F.Root defaultValues={{title:'',tags:[]}} onSubmit={()=>{}}><SettingsGroup.Item field title="Title" description="Help"><F.Input name="title" /></SettingsGroup.Item><SettingsGroup.Item field title="Tags"><F.TagPicker name="tags" options={[]} /></SettingsGroup.Item></F.Root>;
