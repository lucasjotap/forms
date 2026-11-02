import { env } from 'cloudflare:workers';
import { allowed,activities,validToken,isAdmin } from '../../form-security';
export async function POST(req:Request){
 const origin=req.headers.get('origin');if(!origin||origin!==new URL(req.url).origin)return Response.json({error:'Origem inválida.'},{status:403});
 if(!req.headers.get('content-type')?.includes('application/json'))return Response.json({error:'Formato inválido.'},{status:415});
 const text=await req.text();if(text.length>6000)return Response.json({error:'Envio grande demais.'},{status:413});
 let body;try{body=JSON.parse(text)}catch{return Response.json({error:'Envio inválido.'},{status:400})}
 if(!body||body.consent!==true||!await validToken(body.token))return Response.json({error:'Atualize a página e confirme o envio das respostas.'},{status:403});
 if(body.website)return Response.json({error:'Envio inválido.'},{status:400});
 if(typeof body.id!=='string'||!/^[0-9a-f-]{36}$/.test(body.id))return Response.json({error:'Identificador inválido.'},{status:400});
 const a=body.answers;if(!a||typeof a!=='object'||Array.isArray(a))return Response.json({error:'Respostas ausentes.'},{status:400});
 for(const [k,opts]of Object.entries(allowed))if(!opts.includes(a[k]))return Response.json({error:'Responda as oito perguntas.'},{status:400});
 if(typeof a.business!=='string'||a.business.length>100||!activities.includes(a.activity))return Response.json({error:'Dados inválidos.'},{status:400});
 const answers=Object.fromEntries(Object.keys(allowed).map(k=>[k,a[k]]));
 if(!env.DB)return Response.json({error:'Coleta temporariamente indisponível.'},{status:503});
 await env.DB.prepare('INSERT INTO responses (id,received_at,business,activity,answers) VALUES (?,?,?,?,?) ON CONFLICT(id) DO NOTHING').bind(body.id,new Date().toISOString(),a.business.trim(),a.activity,JSON.stringify(answers)).run();
 return Response.json({ok:true,id:body.id},{headers:{'Cache-Control':'no-store'}});
}
export async function GET(){
 if(!await isAdmin())return Response.json({error:'Acesso restrito.'},{status:403});
 const result=await env.DB!.prepare('SELECT * FROM responses ORDER BY received_at DESC LIMIT 5000').all();
 return Response.json({responses:result.results},{headers:{'Cache-Control':'no-store','Content-Disposition':'attachment; filename="respostas-oficina.json"'}});
}
