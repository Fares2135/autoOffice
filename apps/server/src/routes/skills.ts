import { Hono } from 'hono';
import { readSkill } from '../skills/index';

export function skillsRouter() {
  const r = new Hono();

  r.get('/:name{.+}', (c) => {
    const name = c.req.param('name');
    const body = readSkill(name);
    if (body == null) return c.json({ error: `Unknown skill '${name}'` }, 404);
    return c.json({ name, body });
  });

  return r;
}
