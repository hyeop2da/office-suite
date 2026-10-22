// 작업 중 변경 사항의 임시 보존(IndexedDB). 이 기기 브라우저에만 저장되고 서버로 보내지 않는다.
export interface Draft { name: string; bytes: Uint8Array; edits: [string, string][]; savedAt: number }

const DB = 'office-suite';
const STORE = 'drafts';
const KEY = 'current';

function open(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => { req.result.createObjectStore(STORE); };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function run<T>(mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await open();
  try {
    return await new Promise<T>((resolve, reject) => {
      const tx = db.transaction(STORE, mode);
      const req = fn(tx.objectStore(STORE));
      tx.oncomplete = () => resolve(req.result);
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  } finally { db.close(); }
}

export async function saveDraft(d: Draft): Promise<boolean> {
  try { await run('readwrite', s => s.put(d, KEY)); return true; } catch { return false; }
}

export async function loadDraft(): Promise<Draft | null> {
  try { return ((await run('readonly', s => s.get(KEY))) as Draft | undefined) ?? null; } catch { return null; }
}

export async function deleteDraft(): Promise<void> {
  try { await run('readwrite', s => s.delete(KEY)); } catch { /* 저장소를 쓸 수 없으면 지울 것도 없다 */ }
}
