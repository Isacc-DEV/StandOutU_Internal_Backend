import fastifyPlugin from 'fastify-plugin';
import jwt from 'jsonwebtoken';
import { FastifyReply } from 'fastify';
import { findUserById } from './db';
import { User } from './types';

const JWT_SECRET = process.env.JWT_SECRET || 'dev-smartwork-jwt-secret';

export function signToken(user: User) {
  const payload = { sub: user.id, role: user.role, email: user.email, name: user.name };
  return jwt.sign(payload, JWT_SECRET, { expiresIn: '1d' });
}

export function verifyToken(token: string) {
  return jwt.verify(token, JWT_SECRET) as { sub: string };
}

declare module 'fastify' {
  interface FastifyRequest {
    authUser?: User;
  }
}

export const authGuard = fastifyPlugin(async (instance) => {
  instance.addHook('preHandler', async (request, reply) => {
    const routeUrl = request.routeOptions?.url;
    if (
      routeUrl === '/health' ||
      routeUrl === '/auth/login' ||
      routeUrl === '/auth/signup' ||
      routeUrl === '/calendar/oauth/callback' ||
      (routeUrl && routeUrl.startsWith('/ws/browser')) ||
      (routeUrl && routeUrl.startsWith('/ws/community')) ||
      (routeUrl && routeUrl.startsWith('/ws/notifications'))
    ) {
      return;
    }
    // Also allow websocket upgrade paths detected via raw url.
    if (
      request.raw?.url?.startsWith('/ws/browser') ||
      request.raw?.url?.startsWith('/ws/community') ||
      request.raw?.url?.startsWith('/ws/notifications')
    ) {
      return;
    }
    const header = request.headers.authorization;
    if (!header || !header.startsWith('Bearer ')) {
      return reply.status(401).send({ message: 'Missing token' });
    }
    const token = header.slice('Bearer '.length);
    try {
      const decoded = jwt.verify(token, JWT_SECRET) as { sub: string };
      const user = await findUserById(decoded.sub);
      if (!user) {
        return reply.status(401).send({ message: 'Invalid token' });
      }
      request.authUser = user;
    } catch {
      return reply.status(401).send({ message: 'Invalid token' });
    }
  });
});

export function forbidObserver(reply: FastifyReply, user?: User) {
  if (user?.role === 'OBSERVER') {
    reply.status(403).send({ message: 'Observer role cannot access this resource' });
    return true;
  }
  return false;
}
