import { issueToken } from '../../form-security';
export async function GET(){return Response.json({token:await issueToken()},{headers:{'Cache-Control':'no-store'}});}
