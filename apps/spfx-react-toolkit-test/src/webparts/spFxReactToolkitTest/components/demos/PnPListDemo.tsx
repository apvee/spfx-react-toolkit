import { DemoCard } from '../shared/DemoCard';
import { useDemoStyles } from '../shared/demoStyles';
import { DemoIcon } from '../shared/DemoIcon';
import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, Label, MessageBar, MessageBarActions, MessageBarBody } from '@fluentui/react-components';
import * as React from 'react';
import { useSPFxPnPList } from '@apvee/spfx-react-toolkit';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
/**
 * Example 7: useSPFxPnPList - CRUD Operations
 */
export const PnPListDemo: React.FC<{
    readonly selector?: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const demoStyles = useDemoStyles();
    const [writing, setWriting] = React.useState(false);
    const mounted = React.useRef(true);
    React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
    const [listTitle, setListTitle] = React.useState('');
    const [newTitle, setNewTitle] = React.useState('');
    const [editingId, setEditingId] = React.useState<number | null>(null);
    const [editTitle, setEditTitle] = React.useState('');
    const [deleteId, setDeleteId] = React.useState<number>();
    const deleteRequest = React.useRef<{
        readonly id: number;
        readonly trigger: HTMLButtonElement;
        readonly remove: (id: number) => Promise<unknown>;
    }>();
    // React 17 native ref props require the legacy null-initialized RefObject.
    const listInput = React.useRef<HTMLInputElement>(null);
    const restoreDeleteFocus = React.useCallback((trigger: HTMLButtonElement) => {
        if (trigger.isConnected)
            trigger.focus();
        else
            listInput.current?.focus();
    }, []);
    const cancelDelete = React.useCallback(() => {
        const request = deleteRequest.current;
        deleteRequest.current = undefined;
        setDeleteId(undefined);
        if (request)
            restoreDeleteFocus(request.trigger);
    }, [restoreDeleteFocus]);
    const { query, items, loading, error, isEmpty, hasMore, loadMore, clearError, create, update, remove, } = useSPFxPnPList<{
        Id: number;
        Title: string;
    }>(listTitle, { pageSize: 10 });
    const handleLoadList = React.useCallback(() => {
        query(q => q.select('Id', 'Title').orderBy('Id', false)).catch(err => console.error('Load error:', err));
    }, [query]);
    const handleCreate = React.useCallback(async () => {
        if (!newTitle || !listTitle)
            return;
        setWriting(true);
        try {
            await create({ Title: newTitle });
            setNewTitle('');
        }
        catch (err) {
            console.error('Create error:', err);
        }
        finally {
            if (mounted.current)
                setWriting(false);
        }
    }, [create, newTitle, listTitle]);
    const handleUpdate = React.useCallback(async (id: number) => {
        if (!editTitle)
            return;
        setWriting(true);
        try {
            await update(id, { Title: editTitle });
            setEditingId(null);
            setEditTitle('');
        }
        catch (err) {
            console.error('Update error:', err);
        }
        finally {
            if (mounted.current)
                setWriting(false);
        }
    }, [update, editTitle]);
    const requestDelete = React.useCallback((id: number, trigger: HTMLButtonElement) => {
        // Capture the intended operation when consent opens; a later render cannot retarget it.
        deleteRequest.current = { id, trigger, remove };
        setDeleteId(id);
    }, [remove]);
    const approveDelete = React.useCallback(async () => {
        const request = deleteRequest.current;
        if (!request)
            return;
        // Consume this approval before dispatch, including rapid repeated activation.
        deleteRequest.current = undefined;
        setDeleteId(undefined);
        // Refetch publishes rows after removal resolves; focus a stable control before dispatch.
        listInput.current?.focus();
        setWriting(true);
        try {
            await request.remove(request.id);
        }
        catch (err) {
            console.error('Delete error:', err);
        }
        finally {
            if (mounted.current)
                setWriting(false);
        }
        // Completion/refetch leave any subsequent user-selected focus untouched.
    }, []);
    return (<><DemoCard title="Example 3: useSPFxPnPList - CRUD Operations" description="Load, create, edit and delete items in a selected list." iconName="BulletedList" controls={<>{selector?.(writing || deleteId !== undefined)}
    <div className={demoStyles.actions}>
      <Field label={"List Title"}>
        <Input readOnly={writing || deleteId !== undefined} ref={listInput} value={listTitle} onChange={(_, data) => { if (!writing && deleteId === undefined)
            setListTitle(data.value ?? ''); }} placeholder="e.g., Site Pages"/>
      </Field>

    </div>
    {listTitle && <Input aria-label={"New item title..."} value={newTitle} onChange={(_, data) => setNewTitle(data.value ?? '')} placeholder="New item title..."/>}
  </>} contentPlacement="configuration" commands={<>
    <Button appearance="primary" onClick={handleLoadList} disabled={!listTitle || loading || writing}>{"Load"}</Button>
    {listTitle && <Button appearance="primary" onClick={handleCreate} disabled={!newTitle || loading || writing}>{"Create"}</Button>}
  </>} diagnostics={<>
    <Label>
      <DemoIcon name={"InfoSolid"}/>
      Complete CRUD with auto-refetch. Try with &quot;Site Pages&quot; or &quot;Documents&quot;.
    </Label>
  </>}>

    {error && (<MessageBar>
      <MessageBarBody>
        {error.message}
      </MessageBarBody>
      <MessageBarActions>
        <Button aria-label="Dismiss message" onClick={clearError}>Dismiss</Button>
      </MessageBarActions>
    </MessageBar>)}

    {loading && <MessageBar>
      <MessageBarBody>Loading...</MessageBarBody>
    </MessageBar>}

    {isEmpty && !loading && !error && <MessageBar>
      <MessageBarBody>No items found.</MessageBarBody>
    </MessageBar>}

    {!loading && !error && items.length > 0 && (<div className={demoStyles.panel}>
      <Label>Items ({items.length}):</Label>
      {items.map(item => (<div className={demoStyles.actions} key={item.Id}>
        {editingId === item.Id ? (<>
          <Input aria-label={"Value"} value={editTitle} onChange={(_, data) => setEditTitle(data.value ?? '')}/>
          <Button disabled={writing} onClick={() => handleUpdate(item.Id)}>{"Save"}</Button>
          <Button onClick={() => { setEditingId(null); setEditTitle(''); }}>{"Cancel"}</Button>
        </>) : (<>
          <div className={demoStyles.observations}>#{item.Id} - {item.Title}</div>
          <Button onClick={() => { setEditingId(item.Id); setEditTitle(item.Title); }}>{"Edit"}</Button>
          <Button disabled={writing} onClick={event => requestDelete(item.Id, event.currentTarget)}>{"Delete"}</Button>
        </>)}
      </div>))}
      {hasMore && <Button appearance="primary" onClick={loadMore} disabled={loading}>{"Load More"}</Button>}
    </div>)}

  </DemoCard>
    <Dialog open={deleteId !== undefined} onOpenChange={(_, data) => { if (!data.open)
        cancelDelete(); }}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Delete this item?</DialogTitle>
          <DialogContent>Delete item #{deleteId}? This action cannot be undone.</DialogContent>
          <DialogActions>
            <Button onClick={cancelDelete}>Cancel</Button>
            <Button appearance="primary" onClick={approveDelete}>Delete item</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog></>);
};
