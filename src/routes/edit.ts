import type { FastifyInstance } from 'fastify';
import { recordAudit } from '../audit.js';
import { joinEditSession } from '../collab/room.js';
import { InvalidTokenError, verifyViewToken } from '../security.js';
import { store } from '../storage.js';

interface EditQuery {
  readonly token?: string;
}

const editQuerySchema = {
  type: 'object',
  properties: { token: { type: 'string' } },
} as const;

/**
 * docx 실시간 동시편집 WebSocket 엔드포인트.
 * 열람 토큰과 동일한 JWT를 쓰되 perms.canEdit이 있어야 하며, 검증은 여기서 직접 한다
 * (viewer.ts의 resolveView는 HTTP FastifyReply 기반이라 WS 업그레이드 이후에는 재사용할 수 없다 —
 * 업그레이드가 이미 끝난 뒤에는 HTTP 에러 응답 대신 WS close 코드로 거부를 알린다).
 */
export async function editRoutes(app: FastifyInstance): Promise<void> {
  // 경로의 :docId는 y-websocket 클라이언트(WebsocketProvider)가 요구하는 room 경로 형태를
  // 맞추기 위한 것일 뿐 신뢰하지 않는다 — 실제 문서 식별·권한은 항상 token(JWT)의 docId를 쓴다.
  app.get<{ Querystring: EditQuery; Params: { docId: string } }>(
    '/ws/edit/:docId',
    { websocket: true, schema: { querystring: editQuerySchema } },
    async (socket, request) => {
      // 소켓은 이 핸들러가 시작되는 시점에 이미 열려 있고 클라이언트는 열리자마자 syncStep1을
      // 보낸다 — 그런데 아래 인증·조회는 전부 비동기(await)라 메시지 리스너를 그 뒤에 붙이면
      // 그 사이 도착한 메시지가 리스너 없는 EventEmitter에서 그냥 유실된다(실제로 재현된 버그:
      // 클라이언트의 첫 syncStep1이 사라져 동기화가 끝나지 않았다). 그래서 인증 성공 여부와
      // 무관하게 가장 먼저 메시지를 큐에 담아두고, 성공하면 그 큐를 joinEditSession에 넘긴다.
      const earlyMessages: Array<ArrayBuffer | Buffer> = [];
      const queueEarly = (data: ArrayBuffer | Buffer): void => {
        earlyMessages.push(data);
      };
      socket.on('message', queueEarly);

      const { token } = request.query;
      if (!token) {
        socket.off('message', queueEarly);
        socket.close(4401, '열람 토큰이 없습니다.');
        return;
      }

      let claims;
      try {
        claims = verifyViewToken(token);
      } catch (error) {
        socket.off('message', queueEarly);
        if (error instanceof InvalidTokenError) {
          socket.close(4401, error.message);
          return;
        }
        throw error;
      }

      if (!claims.perms.canEdit) {
        socket.off('message', queueEarly);
        await recordAudit({ type: 'edit_denied', docId: claims.docId, viewer: claims.viewer });
        socket.close(4403, '편집 권한이 없습니다.');
        return;
      }

      const meta = await store.readMeta(claims.docId);
      if (!meta || meta.status !== 'ready') {
        socket.off('message', queueEarly);
        socket.close(4404, '문서를 찾을 수 없습니다.');
        return;
      }

      await recordAudit({ type: 'edit_joined', docId: claims.docId, viewer: claims.viewer });
      socket.off('message', queueEarly);
      await joinEditSession(socket, request.raw, claims.docId, meta.ext, earlyMessages);
    },
  );
}
