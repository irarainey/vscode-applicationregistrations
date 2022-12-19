import * as cp from "child_process";

export const execShell = (cmd: string) =>
new Promise<string>((resolve, reject) => {
    cp.exec(cmd, (err, out) => {
        if (err) {
            return reject(err);
        }
        return resolve(out);
    });
});
