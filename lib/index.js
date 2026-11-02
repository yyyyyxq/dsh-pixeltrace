/**
 * Host half of the dsh-pixeltrace bundle.
 *
 * The entire product lives in the Client module (`lib/client.js`): the Web UI panel
 * that decodes an image, rebuilds it out of PowerPoint objects and downloads the
 * .pptx. The Host half only has to exist so the bundle row can mount.
 */
export const name = 'dsh-pixeltrace';

export function apply() {}
