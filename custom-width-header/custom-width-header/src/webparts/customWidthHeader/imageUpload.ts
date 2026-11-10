import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/** Uploaded header images go here, inside the site's Site Assets library. */
const FOLDER_NAME = 'CustomWidthHeader';

export const IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp'];
const MAX_BYTES = 10 * 1024 * 1024;

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

async function ensureOk(response: SPHttpClientResponse, step: string): Promise<SPHttpClientResponse> {
  if (response.ok) return response;
  let detail = `${response.status} ${response.statusText}`;
  try {
    const body = await response.json();
    detail = body?.error?.message ?? body?.['odata.error']?.message?.value ?? detail;
  } catch {
    // keep the status text
  }
  throw new Error(`${step}: ${detail}`);
}

/**
 * Saves a file chosen with the property pane's "Upload" tab into
 * Site Assets/CustomWidthHeader and returns its server-relative URL.
 *
 * The PnP file picker leaves uploads to the web part, so this makes the REST
 * calls involved: make sure Site Assets exists, look up its folder, make sure
 * our subfolder exists, then add the file under a unique name so two authors
 * uploading "logo.png" never overwrite each other.
 */
export async function uploadImage(
  client: SPHttpClient,
  webUrl: string,
  file: Blob,
  fileName: string
): Promise<string> {
  const extension = (/\.[^.]+$/.exec(fileName)?.[0] ?? '').toLowerCase();
  if (IMAGE_EXTENSIONS.indexOf(extension) < 0) {
    throw new Error(`only ${IMAGE_EXTENSIONS.join(', ')} images can be used`);
  }
  if (file.size > MAX_BYTES) {
    throw new Error('images must be 10 MB or smaller');
  }

  const api = `${webUrl.replace(/\/$/, '')}/_api`;

  const listResponse = await ensureOk(
    await client.post(`${api}/web/lists/EnsureSiteAssetsLibrary`, SPHttpClient.configurations.v1, {}),
    'Opening Site Assets'
  );
  const list = await listResponse.json();

  const rootResponse = await ensureOk(
    await client.get(
      `${api}/web/lists(guid'${list.Id}')/RootFolder?$select=ServerRelativeUrl`,
      SPHttpClient.configurations.v1
    ),
    'Reading Site Assets'
  );
  const root: string = (await rootResponse.json()).ServerRelativeUrl;
  const folder = `${root}/${FOLDER_NAME}`;

  // Adding a folder that already exists returns it, so this is safe to repeat.
  const folderResponse = await client.post(
    `${api}/web/folders/add('${encodeURIComponent(escapeODataString(folder))}')`,
    SPHttpClient.configurations.v1,
    {}
  );
  const target = folderResponse.ok ? folder : root;

  const safeName = fileName.replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '') || 'image';
  const uniqueName = `${Date.now()}-${safeName}`;

  const uploadResponse = await ensureOk(
    await client.post(
      `${api}/web/GetFolderByServerRelativeUrl('${encodeURIComponent(escapeODataString(target))}')` +
        `/Files/add(url='${encodeURIComponent(escapeODataString(uniqueName))}',overwrite=false)`,
      SPHttpClient.configurations.v1,
      { body: await file.arrayBuffer() }
    ),
    'Uploading'
  );
  const uploaded = await uploadResponse.json();
  return uploaded.ServerRelativeUrl as string;
}
