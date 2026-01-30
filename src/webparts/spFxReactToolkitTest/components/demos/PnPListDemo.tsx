import * as React from 'react';
import {
  Stack,
  TextField,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  Separator,
  Label,
  Icon,
} from '@fluentui/react';
import { useSPFxPnPList } from '../../../../hooks';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

/**
 * Example 7: useSPFxPnPList - CRUD Operations
 */
export const PnPListDemo: React.FC = () => {
  const [listTitle, setListTitle] = React.useState('');
  const [newTitle, setNewTitle] = React.useState('');
  const [editingId, setEditingId] = React.useState<number | null>(null);
  const [editTitle, setEditTitle] = React.useState('');

  const {
    query,
    items,
    loading,
    error,
    isEmpty,
    hasMore,
    loadMore,
    clearError,
    create,
    update,
    remove,
  } = useSPFxPnPList<{ Id: number; Title: string }>(listTitle, { pageSize: 10 });

  const handleLoadList = React.useCallback(() => {
    query(q => q.select('Id', 'Title').orderBy('Id', false)).catch(err => console.error('Load error:', err));
  }, [query]);

  const handleCreate = React.useCallback(async () => {
    if (!newTitle || !listTitle) return;
    try {
      await create({ Title: newTitle });
      setNewTitle('');
    } catch (err) {
      console.error('Create error:', err);
    }
  }, [create, newTitle, listTitle]);

  const handleUpdate = React.useCallback(async (id: number) => {
    if (!editTitle) return;
    try {
      await update(id, { Title: editTitle });
      setEditingId(null);
      setEditTitle('');
    } catch (err) {
      console.error('Update error:', err);
    }
  }, [update, editTitle]);

  const handleDelete = React.useCallback(async (id: number) => {
    if (!confirm('Delete this item?')) return;
    try {
      await remove(id);
    } catch (err) {
      console.error('Delete error:', err);
    }
  }, [remove]);

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px', border: '1px solid #edebe9', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="BulletedList" style={{ marginRight: '8px' }} />
        Example 3: useSPFxPnPList - CRUD Operations
      </h3>
      <Separator />

      <Stack horizontal tokens={{ childrenGap: 6 }}>
        <TextField
          label="List Title"
          value={listTitle}
          onChange={(_, newValue) => setListTitle(newValue ?? '')}
          placeholder="e.g., Site Pages"
          styles={{ root: { flexGrow: 1 } }}
        />
        <PrimaryButton
          text="Load"
          onClick={handleLoadList}
          disabled={!listTitle || loading}
          styles={{ root: { marginTop: '28px' } }}
        />
      </Stack>

      {error && (
        <MessageBar onDismiss={clearError}>
          {error.message}
        </MessageBar>
      )}

      {listTitle && (
        <Stack horizontal tokens={{ childrenGap: 6 }} styles={{ root: { padding: '8px', backgroundColor: '#f3f2f1', borderRadius: '4px' } }}>
          <TextField
            value={newTitle}
            onChange={(_, newValue) => setNewTitle(newValue ?? '')}
            placeholder="New item title..."
            styles={{ root: { flexGrow: 1 } }}
          />
          <PrimaryButton text="Create" onClick={handleCreate} disabled={!newTitle || loading} />
        </Stack>
      )}

      {loading && <MessageBar>Loading...</MessageBar>}
      {isEmpty && !loading && <MessageBar>No items found.</MessageBar>}

      {items.length > 0 && (
        <Stack tokens={{ childrenGap: 3 }}>
          <Label>Items ({items.length}):</Label>
          {items.map(item => (
            <Stack key={item.Id} horizontal tokens={{ childrenGap: 6 }} styles={{ root: { padding: '6px', backgroundColor: '#faf9f8', borderRadius: '4px' } }}>
              {editingId === item.Id ? (
                <>
                  <TextField value={editTitle} onChange={(_, v) => setEditTitle(v ?? '')} styles={{ root: { flexGrow: 1 } }} />
                  <DefaultButton text="Save" onClick={() => handleUpdate(item.Id)} />
                  <DefaultButton text="Cancel" onClick={() => { setEditingId(null); setEditTitle(''); }} />
                </>
              ) : (
                <>
                  <div style={{ flexGrow: 1 }}>#{item.Id} - {item.Title}</div>
                  <DefaultButton text="Edit" onClick={() => { setEditingId(item.Id); setEditTitle(item.Title); }} />
                  <DefaultButton text="Delete" onClick={() => handleDelete(item.Id)} />
                </>
              )}
            </Stack>
          ))}
          {hasMore && <PrimaryButton text="Load More" onClick={loadMore} disabled={loading} />}
        </Stack>
      )}

      <Label>
        <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
        Complete CRUD with auto-refetch. Try with &quot;Site Pages&quot; or &quot;Documents&quot;.
      </Label>
    </Stack>
  );
};
