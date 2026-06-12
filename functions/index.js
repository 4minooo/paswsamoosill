const { setGlobalOptions } = require("firebase-functions/v2");
const { HttpsError, onCall } = require("firebase-functions/v2/https");
const admin = require("firebase-admin");

admin.initializeApp();
setGlobalOptions({ region: "asia-northeast3" });

const WORKSPACE_ID = "social-welfare-office";
const DEFAULT_ADMIN_STUDENT_IDS = new Set(["202210276"]);

exports.updateMemberPassword = onCall(async (request) => {
  const callerStudentId = getStudentIdFromAuth(request.auth);
  if (!callerStudentId) {
    throw new HttpsError("unauthenticated", "Login is required.");
  }

  const callerIsAdmin = await isApprovedAdmin(callerStudentId);
  if (!callerIsAdmin) {
    throw new HttpsError("permission-denied", "Admin permission is required.");
  }

  const studentId = normalizeStudentId(request.data?.studentId);
  const newPassword = String(request.data?.newPassword || "");
  if (!studentId || newPassword.length < 6) {
    throw new HttpsError("invalid-argument", "Password must be at least 6 characters.");
  }

  const email = `${studentId}@pknu-work.app`;
  let userRecord;
  try {
    userRecord = await admin.auth().getUserByEmail(email);
  } catch (error) {
    throw new HttpsError("not-found", "Target user was not found.");
  }

  await admin.auth().updateUser(userRecord.uid, {
    password: newPassword
  });

  await admin.firestore().doc(`users/${studentId}`).set(
    {
      passwordChangedAt: admin.firestore.FieldValue.serverTimestamp(),
      passwordChangedBy: callerStudentId
    },
    { merge: true }
  );

  return { ok: true };
});

async function isApprovedAdmin(studentId) {
  if (DEFAULT_ADMIN_STUDENT_IDS.has(studentId)) {
    return true;
  }

  const workspaceSnap = await admin.firestore().doc(`workspaces/${WORKSPACE_ID}`).get();
  const members = workspaceSnap.data()?.state?.members || [];
  return members.some((member) => {
    return (
      normalizeStudentId(member.studentId) === studentId &&
      member.status === "approved" &&
      member.role === "admin"
    );
  });
}

function getStudentIdFromAuth(auth) {
  const email = auth?.token?.email || "";
  return normalizeStudentId(email.split("@")[0] || "");
}

function normalizeStudentId(value) {
  return String(value || "").replace(/\s+/g, "").toLowerCase();
}
