import * as React from 'react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import styles from './PropertyPaneControls.module.scss';
import { checkImageUrl } from '../common/url';
import { ACCEPT_ATTRIBUTE, validateImageFile } from '../common/imageFiles';

export interface IImagePickerProps {
  label: string;
  /** Current image address ('' for none). */
  value: string;
  onChange: (url: string) => void;
  /** Uploads the file and resolves to its address. */
  upload: (file: File) => Promise<string>;
}

type Status = { kind: 'idle' } | { kind: 'uploading'; name: string } | { kind: 'done'; name: string } | { kind: 'error'; message: string };

/**
 * Choose a card image: upload one from the computer (saved to the site's
 * Site Assets library) or paste a link to an existing image.
 */
export default function ImagePicker(props: IImagePickerProps): React.ReactElement {
  const fileInput = React.useRef<HTMLInputElement | undefined>(undefined);
  const [status, setStatus] = React.useState<Status>({ kind: 'idle' });
  // Tracked locally so the preview updates at once, without the pane having
  // to rebuild (which would interrupt the upload announcement).
  const [value, setValue] = React.useState(props.value);
  const [linkText, setLinkText] = React.useState(props.value);
  const [previewFailed, setPreviewFailed] = React.useState(false);
  const mounted = React.useRef(true);

  React.useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
    };
  }, []);

  React.useEffect(() => {
    setValue(props.value);
    setLinkText(props.value);
    setPreviewFailed(false);
  }, [props.value]);

  const commit = (url: string): void => {
    setValue(url);
    setPreviewFailed(false);
    props.onChange(url);
  };

  const linkCheck = checkImageUrl(linkText);
  const linkError = linkCheck.ok ? '' : linkCheck.reason;
  const uploading = status.kind === 'uploading';

  const onFileChosen = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = event.target.files && event.target.files[0];
    event.target.value = ''; // let the same file be chosen again later
    if (!file) {
      return;
    }
    const problem = validateImageFile(file);
    if (problem) {
      setStatus({ kind: 'error', message: problem });
      return;
    }
    setStatus({ kind: 'uploading', name: file.name });
    try {
      const url = await props.upload(file);
      if (!mounted.current) {
        return;
      }
      setStatus({ kind: 'done', name: file.name });
      setLinkText(url);
      commit(url);
    } catch (error) {
      if (mounted.current) {
        setStatus({ kind: 'error', message: error instanceof Error ? error.message : String(error) });
      }
    }
  };

  return (
    <fieldset className={styles.fieldset}>
      <legend className={styles.legend}>{props.label}</legend>

      {value && (
        <div className={styles.imagePreview}>
          {previewFailed ? (
            <span className={styles.helpText}>This image couldn&apos;t be loaded. Check the link or upload it again.</span>
          ) : (
            // Decorative here: the address is shown as text right next to it.
            <img src={value} alt="" onError={() => setPreviewFailed(true)} />
          )}
          <span className={styles.imagePreviewName}>{decodeName(value)}</span>
        </div>
      )}

      <div className={styles.buttonRow}>
        <DefaultButton
          text={uploading ? 'Uploading…' : value ? 'Upload a different image' : 'Upload image'}
          disabled={uploading}
          onClick={() => fileInput.current && fileInput.current.click()}
        />
        {value && !uploading && (
          <DefaultButton
            text="Remove image"
            onClick={() => {
              setLinkText('');
              setStatus({ kind: 'idle' });
              commit('');
            }}
          />
        )}
      </div>
      <input
        ref={(element) => {
          fileInput.current = element || undefined;
        }}
        type="file"
        accept={ACCEPT_ATTRIBUTE}
        className={styles.visuallyHiddenInput}
        tabIndex={-1}
        aria-hidden="true"
        onChange={onFileChosen}
      />

      <div role="status" aria-live="polite" className={styles.helpText}>
        {status.kind === 'uploading' && `Uploading ${status.name}…`}
        {status.kind === 'done' && `Uploaded ${status.name} to Site Assets.`}
      </div>
      {status.kind === 'error' && (
        <div role="alert" className={styles.errorText}>
          {status.message}
        </div>
      )}

      <p className={styles.helpText}>PNG, JPG, GIF, WebP or SVG, up to 5 MB. Uploads are saved in this site&apos;s Site Assets library.</p>

      <TextField
        label="Or paste an image link"
        value={linkText}
        placeholder="https://…/image.png"
        onChange={(_event, newValue) => {
          const text = newValue || '';
          setLinkText(text);
          if (checkImageUrl(text).ok) {
            commit(text.trim());
          }
        }}
        errorMessage={linkError}
        autoComplete="off"
        spellCheck={false}
      />
    </fieldset>
  );
}

/** "/sites/hr/SiteAssets/My%20Logo.png" → "My Logo.png" for display. */
function decodeName(url: string): string {
  const last = url.split(/[?#]/)[0].split('/').pop() || url;
  try {
    return decodeURIComponent(last);
  } catch {
    return last;
  }
}
