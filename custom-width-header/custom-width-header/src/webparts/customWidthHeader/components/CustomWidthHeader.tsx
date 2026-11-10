import * as React from 'react';
import { ArrowLeft } from 'lucide-react';
import * as strings from 'CustomWidthHeaderWebPartStrings';
import styles from './CustomWidthHeader.module.scss';
import { getIcon } from '../icons';
import { safeHref, safeImageSrc } from '../urls';
import type { ICustomWidthHeaderWebPartProps } from '../models';
import type { CssVars } from '../theme';

export interface ICustomWidthHeaderProps extends ICustomWidthHeaderWebPartProps {
  isEditMode: boolean;
  /** Used when the back link has no address of its own. */
  siteHomeUrl: string;
  themeVars: CssVars;
  bodyFont: string;
  titleFont: string;
  /** The level actually used, after resolving 'auto' against the page. */
  resolvedHeadingLevel: '1' | '2';
}

const LAYOUT_CLASS = { page: styles.layoutPage, resource: styles.layoutResource };
const HEIGHT_CLASS = { auto: '', medium: styles.heightMedium, large: styles.heightLarge };
const POSITION_CLASS = { center: styles.posCenter, top: styles.posTop, bottom: styles.posBottom };

/**
 * The hero banner from the Manager Hub's lifecycle pages and resource detail
 * view, as one configurable header: back link, small label, icon or image
 * tile, title and description, on a colour or a photo.
 *
 * Images are purely decorative (empty alt, hidden from assistive tech); the
 * text carries all the meaning, so nothing is lost if an image fails to load.
 */
export function CustomWidthHeader(props: ICustomWidthHeaderProps): JSX.Element {
  const { isEditMode, themeVars } = props;
  const Heading = props.resolvedHeadingLevel === '2' ? 'h2' : 'h1';
  const title = props.title.trim();
  const description = props.description.trim();
  const eyebrow = props.eyebrow.trim();
  const backText = props.backLinkText.trim();

  const backHref = safeHref(props.backLinkUrl.trim() || props.siteHomeUrl);
  const showBack = props.showBackLink && backText !== '' && !!backHref;

  const isImage = props.backgroundType === 'image';
  const imageSrc = isImage ? safeImageSrc(props.backgroundImageUrl) : undefined;
  const isPanel = isImage && props.imageTextStyle === 'panel';

  const iconImage = props.iconSource === 'image' ? safeImageSrc(props.iconImageUrl) : undefined;
  const Icon = getIcon(props.iconName);
  const showIcon = props.iconSource === 'icon' || (props.iconSource === 'image' && (!!iconImage || isEditMode));

  const rootClass = [
    styles.header,
    LAYOUT_CLASS[props.layout] ?? styles.layoutPage,
    HEIGHT_CLASS[props.height] ?? '',
    isImage ? styles.hasImage : '',
    isPanel ? styles.panelMode : ''
  ]
    .filter(Boolean)
    .join(' ');

  return (
    <header
      className={rootClass}
      style={{ ...themeVars, '--ph-font': props.bodyFont, '--ph-title-font': props.titleFont } as React.CSSProperties}
    >
      {imageSrc && (
        <img
          className={[styles.backgroundImage, POSITION_CLASS[props.imagePosition] ?? ''].join(' ')}
          src={imageSrc}
          alt=""
          aria-hidden="true"
        />
      )}
      {isImage && !isPanel && <div className={styles.overlay} aria-hidden="true" />}

      <div className={styles.inner}>
        <div className={styles.content}>
          {showBack && (
            <a
              className={styles.backLink}
              href={backHref}
              data-interception="on"
              onClick={isEditMode ? (event) => event.preventDefault() : undefined}
            >
              <ArrowLeft aria-hidden="true" focusable="false" />
              {backText}
            </a>
          )}

          {eyebrow && <p className={styles.eyebrow}>{eyebrow}</p>}

          <div className={styles.row}>
            {showIcon && (
              <span className={styles.iconTile} aria-hidden="true">
                {props.iconSource === 'image' ? (
                  iconImage ? <img className={styles.iconImage} src={iconImage} alt="" /> : null
                ) : (
                  <Icon aria-hidden="true" focusable="false" />
                )}
              </span>
            )}
            <div className={styles.text}>
              {(title || isEditMode) && (
                <Heading className={[styles.title, title ? '' : styles.placeholder].join(' ')}>
                  {title || strings.TitlePlaceholder}
                </Heading>
              )}
              {description && <p className={styles.description}>{description}</p>}
            </div>
          </div>

          {isEditMode && isImage && !imageSrc && <p className={styles.editHint}>{strings.NoImageYet}</p>}
        </div>
      </div>
    </header>
  );
}
