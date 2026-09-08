const SITE_URL='https://www.marqcleanai.vertexsg.co.za';
const TITLE='MarqClean AI — Clean data. Clear decisions.';
const DESCRIPTION='MarqClean AI is a browser-first data operations platform for cleaning, validating, transforming and reconciling Excel, CSV and PDF data. Built for finance, operations, compliance, marketing and data teams.';
const OG_IMAGE=`${SITE_URL}/brand-logo.svg`;

function upsertMeta(name:string,content:string,property=false){
  const attr=property?'property':'name';
  let el=document.head.querySelector<HTMLMetaElement>(`meta[${attr}="${name}"]`);
  if(!el){el=document.createElement('meta');el.setAttribute(attr,name);document.head.appendChild(el)}
  el.setAttribute('content',content);
}
function upsertLink(rel:string,href:string,type?:string){
  let el=document.head.querySelector<HTMLLinkElement>(`link[rel="${rel}"]`);
  if(!el){el=document.createElement('link');el.rel=rel;document.head.appendChild(el)}
  el.href=href;if(type)el.type=type;
}

export function installBrandSeo(){
  document.title=TITLE;
  upsertMeta('description',DESCRIPTION);
  upsertMeta('theme-color','#001f54');upsertMeta('robots','index, follow');upsertMeta('author','Vertex Stream Technologies');
  upsertMeta('og:site_name','MarqClean AI',true);upsertMeta('og:title',TITLE,true);upsertMeta('og:description',DESCRIPTION,true);upsertMeta('og:type','website',true);upsertMeta('og:url',SITE_URL,true);upsertMeta('og:image',OG_IMAGE,true);upsertMeta('og:locale','en_ZA',true);
  upsertMeta('twitter:card','summary_large_image');upsertMeta('twitter:title',TITLE);upsertMeta('twitter:description',DESCRIPTION);upsertMeta('twitter:image',OG_IMAGE);
  upsertLink('canonical',SITE_URL);upsertLink('icon','/brand-logo.svg','image/svg+xml');upsertLink('apple-touch-icon','/brand-logo.svg','image/svg+xml');
  let schema=document.getElementById('marqclean-schema');
  if(!schema){schema=document.createElement('script');schema.id='marqclean-schema';schema.type='application/ld+json';document.head.appendChild(schema)}
  schema.textContent=JSON.stringify({
    '@context':'https://schema.org','@graph':[
      {'@type':'Organization','@id':`${SITE_URL}/#organization`,'name':'Vertex Stream Technologies','url':SITE_URL,'logo':OG_IMAGE},
      {'@type':'WebSite','@id':`${SITE_URL}/#website`,'name':'MarqClean AI','url':SITE_URL,'publisher':{'@id':`${SITE_URL}/#organization`}},
      {'@type':'SoftwareApplication','name':'MarqClean AI','url':SITE_URL,'applicationCategory':'BusinessApplication','operatingSystem':'Web','isAccessibleForFree':true,'description':DESCRIPTION,'featureList':['CSV and Excel cleaning','Data validation','Spreadsheet automation','Reconciliation','PDF statement conversion','Duplicate detection','Data transformation']}
    ]
  });
}
