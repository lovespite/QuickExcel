import english from './en-US.json';

const useEnglish = new URLSearchParams(location.hash.slice(1)).get('lang') === 'en-US';
const catalog: Readonly<Record<string, string>> = english;

export function uiText(source: string, ...values: (string | number)[]): string {
  const template = useEnglish ? catalog[source] ?? source : source;
  return template.replace(/\{(\d+)\}/g, (_, index) => String(values[Number(index)]));
}

// Only translate marked application controls, before editor or PDF content is attached.
if (useEnglish) {
  document.documentElement.lang = 'en-US';
  for (const element of document.querySelectorAll<HTMLElement>('[data-i18n]')) {
    for (const field of element.dataset.i18n!.split(' ')) {
      if (field === 'text') element.textContent = uiText(element.textContent ?? '');
      else element.setAttribute(field, uiText(element.getAttribute(field) ?? ''));
    }
  }
}
