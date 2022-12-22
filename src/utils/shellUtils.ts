import * as ChildProcess from "child_process";

// This function is used to execute shell commands
export const execShellCmd = (cmd: string) =>
    new Promise<string>((resolve, reject) => {
        ChildProcess.exec(cmd, (error, response) => {
            if (error) {
                return reject(error);
            }
            return resolve(response);
        });
    });
