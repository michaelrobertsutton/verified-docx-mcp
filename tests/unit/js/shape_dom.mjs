// Small namespace-aware XML model for offline inventory fixtures.
export class DOMParser {
  parseFromString(xml) {
    const doc = {children:[], getElementsByTagName(name){return name==='parsererror' && !xml.endsWith('>') ? [{}] : [];},
      getElementsByTagNameNS(ns,name){return descendants(this,ns,name);}};
    const stack=[doc];
    for(const token of xml.match(/<[^>]+>|[^<]+/g)||[]) {
      if(token.startsWith('<?'))continue;
      if(token.startsWith('</')){stack.pop();continue;}
      if(token.startsWith('<')) {
        const match=token.match(/^<([^\s/>]+)/);if(!match)continue;
        const attrs={};for(const m of token.matchAll(/([^\s=]+)="([^"]*)"/g))attrs[m[1]]=m[2];
        const parent=stack.at(-1),namespaces={...parent.namespaces};
        for(const [k,v] of Object.entries(attrs))if(k.startsWith('xmlns:'))namespaces[k.slice(6)]=v;
        const [prefix,localName]=match[1].split(':');
        const node={children:[],parentNode:parent,namespaces,localName:localName||prefix,
          namespaceURI:namespaces[prefix],getAttribute(k){return attrs[k]??null;},
          get textContent(){return this.children.map(c=>typeof c==='string'?c:c.textContent).join('');},
          getElementsByTagNameNS(ns,name){return descendants(this,ns,name);}};
        parent.children.push(node);if(!token.endsWith('/>'))stack.push(node);
      } else stack.at(-1).children.push(token);
    }
    return doc;
  }
}
function descendants(node,ns,name) {
  return (node.children||[]).filter(c=>typeof c!=='string').flatMap(c=>
    [...(c.namespaceURI===ns&&c.localName===name?[c]:[]),...descendants(c,ns,name)]);
}
export const EMPTY='<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>';
