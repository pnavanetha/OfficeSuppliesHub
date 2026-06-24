import 'react-toastify/dist/ReactToastify.css';
import { toast, ToastOptions } from "react-toastify";

export type ToastType = "success" | "error" | "info" | "warning";

const defaultOptions: ToastOptions = {
  position: "top-right",
  autoClose: 3000,
  // autoClose: false,
  hideProgressBar: false,
  closeOnClick: true,
  pauseOnHover: true,
  draggable: true,
  theme: "colored"
}

export const showToast = (type: ToastType, message: string) => {
    toast[type](message, defaultOptions);
};

export const showSuccess = (msg: string) => toast.success(msg, defaultOptions );
export const showError = (msg: string) => toast.error(msg, defaultOptions );
export const showInfo  = (msg: string) => toast.info(msg, defaultOptions );
export const showWarning = (msg: string) => toast.warning(msg, defaultOptions );


