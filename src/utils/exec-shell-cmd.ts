import * as ChildProcess from "child_process";

export const execShellCmd = async (cmd: string) => {
    return new Promise<string>((resolve, reject) => {
        ChildProcess.exec(cmd, (error: any, response: any) => {
            if (error) {
                return reject(error);
            }
            return resolve(response);
        });
    });
};