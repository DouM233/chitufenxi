// 对象存储适配层：基于 coze-coding-dev-sdk 的 S3Storage（S3 兼容）。
// 未配置 COZE_BUCKET_ENDPOINT_URL / COZE_BUCKET_NAME 时 storageEnabled=false，
// 调用方据此降级为本地文件行为，系统功能不受影响。
import { S3Storage } from "coze-coding-dev-sdk";

const endpointUrl =
  process.env.COZE_BUCKET_ENDPOINT_URL || process.env.CHITU_S3_ENDPOINT || "";
const bucketName =
  process.env.COZE_BUCKET_NAME || process.env.CHITU_S3_BUCKET || "";
const accessKey = process.env.CHITU_S3_ACCESS_KEY_ID || "";
const secretKey = process.env.CHITU_S3_SECRET_ACCESS_KEY || "";
const region = process.env.CHITU_S3_REGION || "cn-beijing";

export const storageEnabled = Boolean(endpointUrl && bucketName);

export const JOB_SNAPSHOT_PREFIX = "chitu/jobs/";
export const REPORT_PREFIX = "chitu/reports/";

let clientPromise = null;

function getClient() {
  if (!clientPromise) {
    clientPromise = Promise.resolve(
      new S3Storage({ endpointUrl, accessKey, secretKey, bucketName, region })
    );
  }
  return clientPromise;
}

// 上传对象，返回存储端生成的实际 key（含 UUID 前缀），失败返回 null
export async function putObject(key, buffer, contentType) {
  if (!storageEnabled) return null;
  try {
    const client = await getClient();
    return await client.uploadFile({
      fileContent: buffer,
      fileName: key,
      contentType: contentType || "application/octet-stream",
    });
  } catch (error) {
    console.error(
      `[object-storage] 上传失败 key=${key}:`,
      error && error.message ? error.message : error
    );
    return null;
  }
}

// 生成下载用预签名 URL，失败返回 null
export async function presignGetUrl(key, expireSeconds = 86400) {
  if (!storageEnabled) return null;
  try {
    const client = await getClient();
    return await client.generatePresignedUrl({ key, expireTime: expireSeconds });
  } catch (error) {
    console.error(
      `[object-storage] 生成签名 URL 失败 key=${key}:`,
      error && error.message ? error.message : error
    );
    return null;
  }
}

// 读取对象内容，失败返回 null
export async function readObject(key) {
  if (!storageEnabled) return null;
  try {
    const client = await getClient();
    return await client.readFile({ fileKey: key });
  } catch (error) {
    console.error(
      `[object-storage] 读取失败 key=${key}:`,
      error && error.message ? error.message : error
    );
    return null;
  }
}

// 列举指定前缀下的对象 key，失败返回 []
export async function listKeys(prefix, maxKeys = 1000) {
  if (!storageEnabled) return [];
  try {
    const client = await getClient();
    const result = await client.listFiles({ prefix, maxKeys });
    return Array.isArray(result && result.keys) ? result.keys : [];
  } catch (error) {
    console.error(
      `[object-storage] 列举失败 prefix=${prefix}:`,
      error && error.message ? error.message : error
    );
    return [];
  }
}

// 健康检查输出（不含任何凭证信息）
export function storageInfo() {
  return {
    enabled: storageEnabled,
    bucket: storageEnabled ? bucketName : null,
    endpoint_configured: Boolean(endpointUrl),
  };
}
