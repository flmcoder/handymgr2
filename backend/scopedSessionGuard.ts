export type ScopedSessionGuardDependencies<RequestType, ResponseType, SessionType> = {
  requireSession: (request: RequestType, response: ResponseType) => Promise<SessionType | null>;
  applyScope: (request: RequestType, response: ResponseType, session: SessionType) => boolean;
};

export async function enforceScopedSession<RequestType, ResponseType, SessionType>(
  request: RequestType,
  response: ResponseType,
  dependencies: ScopedSessionGuardDependencies<RequestType, ResponseType, SessionType>,
): Promise<boolean> {
  const session = await dependencies.requireSession(request, response);
  if (!session) return false;
  return dependencies.applyScope(request, response, session);
}