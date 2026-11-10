/**
 * Everything the property pane needs that readers never do. Loaded on demand
 * in loadPropertyPaneResources(), so page visitors don't download it.
 */
import PnPTelemetry from '@pnp/telemetry-js';

// The PnP property-pane controls otherwise report usage to an external endpoint.
PnPTelemetry.getInstance().optOut();

export { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
export { PropertyPaneIconPicker } from './PropertyPaneIconPicker';
